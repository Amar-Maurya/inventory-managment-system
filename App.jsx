import React, { useState, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  Eye, EyeOff, Info, Grid3x3, ArrowRight, X, LayoutDashboard, Package,
  Plus, Trash2, Search, Bell, AlertTriangle, DollarSign, Users, Boxes,
  Layers, Droplet, Download, Upload, FileSpreadsheet, LogOut, Check,
  ChevronDown, Image as ImageIcon,
} from "lucide-react";

const COLORS = {
  graphite: "#1C2B33",
  graphiteDeep: "#131E24",
  porcelain: "#FAFAF9",
  surface: "#FFFFFF",
  teal: "#1F6F78",
  tealLight: "#2C8E97",
  tealFaint: "#EFF6F6",
  sand: "#C08A3E",
  sandFaint: "#FAF2E4",
  rust: "#B04A3B",
  rustFaint: "#FCEEEC",
  ink: "#22262A",
  mist: "#8A9499",
  line: "#E6E4DF",
};

const FONT_IMPORT = `
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@500;600;700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');
`;

const CATEGORY_STYLE = {
  Tiles: { icon: Layers, gradient: "linear-gradient(135deg, #2C8E97, #1C2B33)" },
  Adhesive: { icon: Droplet, gradient: "linear-gradient(135deg, #C08A3E, #8A431F)" },
};

const EXPORT_COLUMNS = ["Name", "Category", "Quantity", "Price", "Supplier", "Low Stock Threshold", "Date Added"];

const INITIAL_PRODUCTS = [
  { id: 1, name: "Carrara Matte 600x600", category: "Tiles", qty: 42, threshold: 50, price: 850, supplier: "Marmo Traders", dateAdded: "2026-07-02" },
  { id: 2, name: "Ceramic Bond Pro-X", category: "Adhesive", qty: 8, threshold: 20, price: 320, supplier: "BondWorks Ltd", dateAdded: "2026-06-18" },
  { id: 3, name: "Slate Grey Textured 300x600", category: "Tiles", qty: 130, threshold: 40, price: 540, supplier: "Northstone Supply", dateAdded: "2026-05-27" },
  { id: 4, name: "Flexi-Set Tile Adhesive 20kg", category: "Adhesive", qty: 5, threshold: 15, price: 410, supplier: "BondWorks Ltd", dateAdded: "2026-07-20" },
  { id: 5, name: "Terracotta Rustic 200x200", category: "Tiles", qty: 76, threshold: 30, price: 380, supplier: "ClayCraft Co", dateAdded: "2026-04-11" },
  { id: 6, name: "Anti-Slip Porcelain 450x450", category: "Tiles", qty: 18, threshold: 25, price: 620, supplier: "Marmo Traders", dateAdded: "2026-08-01" },
];

// Indian Rupee formatter — uses en-IN grouping (e.g. ₹1,00,000)
function formatINR(amount) {
  return `₹${Number(amount).toLocaleString("en-IN", { maximumFractionDigits: 0 })}`;
}

// ---------- helpers ----------
function toRows(products) {
  return products.map((p) => ({
    Name: p.name,
    Category: p.category,
    Quantity: p.qty,
    Price: p.price,
    Supplier: p.supplier,
    "Low Stock Threshold": p.threshold,
    "Date Added": p.dateAdded,
  }));
}

function exportToExcel(products, filename) {
  const ws = XLSX.utils.json_to_sheet(toRows(products), { header: EXPORT_COLUMNS });
  ws["!cols"] = [{ wch: 28 }, { wch: 12 }, { wch: 10 }, { wch: 10 }, { wch: 20 }, { wch: 18 }, { wch: 14 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Inventory");
  XLSX.writeFile(wb, filename);
}

function downloadSampleTemplate() {
  const sample = [
    { Name: "Sample Matte Tile 300x300", Category: "Tiles", Quantity: 50, Price: 450, Supplier: "Sample Supplier Co", "Low Stock Threshold": 10, "Date Added": "2026-08-01" },
  ];
  const ws = XLSX.utils.json_to_sheet(sample, { header: EXPORT_COLUMNS });
  ws["!cols"] = [{ wch: 28 }, { wch: 12 }, { wch: 10 }, { wch: 10 }, { wch: 20 }, { wch: 18 }, { wch: 14 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Template");
  XLSX.writeFile(wb, "groutline-import-template.xlsx");
}

function parseImportedRows(rows) {
  return rows.map((r, i) => ({
    id: Date.now() + i,
    name: r["Name"] || "Untitled product",
    category: r["Category"] === "Adhesive" ? "Adhesive" : "Tiles",
    qty: Number(r["Quantity"]) || 0,
    threshold: Number(r["Low Stock Threshold"]) || 10,
    price: Number(r["Price"]) || 0,
    supplier: r["Supplier"] || "—",
    dateAdded: r["Date Added"] || new Date().toISOString().slice(0, 10),
  }));
}

// ---------- small shared UI ----------
function ProductVisual({ category, size = 40 }) {
  const style = CATEGORY_STYLE[category] || CATEGORY_STYLE.Tiles;
  const Icon = style.icon;
  return (
    <div style={{ background: style.gradient, width: size, height: size }} className="rounded-lg flex items-center justify-center flex-shrink-0">
      <Icon size={size * 0.42} color="#fff" strokeWidth={1.75} />
    </div>
  );
}

function StatCard({ icon: Icon, label, value, accent }) {
  return (
    <div style={{ background: COLORS.surface, border: `1px solid ${COLORS.line}` }} className="rounded-xl p-4 flex items-center gap-3">
      <div style={{ background: accent.faint }} className="w-9 h-9 rounded-lg flex items-center justify-center flex-shrink-0">
        <Icon size={16} color={accent.solid} />
      </div>
      <div className="min-w-0">
        <p style={{ fontFamily: "'Sora', sans-serif" }} className="text-lg font-semibold leading-tight truncate">{value}</p>
        <p style={{ color: COLORS.mist }} className="text-xs">{label}</p>
      </div>
    </div>
  );
}

// ---------- Login ----------
function TileGrid() {
  const tiles = Array.from({ length: 30 });
  return (
    <div style={{ display: "grid", gridTemplateColumns: "repeat(6, 1fr)", gap: "6px", maxWidth: "280px" }} aria-hidden="true">
      {tiles.map((_, i) => {
        const delay = (i % 6) * 0.09 + Math.floor(i / 6) * 0.12;
        const accent = i % 11 === 0;
        return (
          <div key={i} style={{
            aspectRatio: "1 / 1", borderRadius: "3px",
            background: accent ? "linear-gradient(135deg, #2C8E97, #1F6F78)" : "rgba(255,255,255,0.06)",
            border: "1px solid rgba(255,255,255,0.08)", animation: `tileLay 0.5s ease-out ${delay}s both`,
          }} />
        );
      })}
      <style>{`@keyframes tileLay { from { opacity: 0; transform: scale(0.6); } to { opacity: 1; transform: scale(1); } }`}</style>
    </div>
  );
}

function LoginScreen({ onLogin }) {
  const [showPassword, setShowPassword] = useState(false);
  const [showInfo, setShowInfo] = useState(false);
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");

  const handleSubmit = (e) => {
    e.preventDefault();
    onLogin();
  };

  return (
    <div style={{ fontFamily: "'Inter', sans-serif", color: COLORS.ink, background: COLORS.porcelain, height: "100vh", overflowY: "auto" }} className="flex flex-col md:flex-row">
      <style>{FONT_IMPORT}</style>
      <div style={{ background: `linear-gradient(160deg, ${COLORS.graphite}, ${COLORS.graphiteDeep})` }} className="relative w-full md:w-1/2 flex flex-col justify-between px-8 py-10 md:px-14 md:py-14 text-white overflow-hidden">
        <div className="flex items-center gap-3">
          <div style={{ background: COLORS.teal }} className="w-9 h-9 rounded-md flex items-center justify-center flex-shrink-0">
            <Grid3x3 size={18} color="#fff" />
          </div>
          <span style={{ fontFamily: "'Sora', sans-serif" }} className="text-lg font-semibold">GroutLine</span>
        </div>
        <div className="my-10 md:my-0">
          <p style={{ fontFamily: "'IBM Plex Mono', monospace", color: COLORS.tealLight, letterSpacing: "0.08em" }} className="text-xs uppercase mb-4">Inventory, laid right</p>
          <h1 style={{ fontFamily: "'Sora', sans-serif", lineHeight: 1.15 }} className="text-3xl md:text-4xl font-semibold max-w-sm mb-6">Every tile, every adhesive batch, accounted for.</h1>
          <p style={{ color: "rgba(255,255,255,0.65)" }} className="max-w-sm text-sm leading-relaxed mb-10">Track stock, suppliers and pricing across your full catalogue.</p>
          <TileGrid />
        </div>
        <p style={{ color: "rgba(255,255,255,0.35)" }} className="text-xs">© 2026 GroutLine</p>
      </div>

      <div className="w-full md:w-1/2 flex items-center justify-center px-6 py-12 md:px-14">
        <div className="w-full max-w-sm">
          <div className="flex items-start justify-between mb-1">
            <h2 style={{ fontFamily: "'Sora', sans-serif" }} className="text-2xl font-semibold">Sign in</h2>
            <button onClick={() => setShowInfo(!showInfo)} style={{ color: COLORS.mist }} className="mt-1 hover:opacity-70" aria-label="Info">
              <Info size={19} />
            </button>
          </div>
          <p style={{ color: COLORS.mist }} className="text-sm mb-6">Enter your details to access the dashboard.</p>

          {showInfo && (
            <div style={{ background: "#EFF6F6", border: `1px solid ${COLORS.teal}22` }} className="rounded-lg p-4 mb-6 text-sm relative">
              <button onClick={() => setShowInfo(false)} className="absolute top-3 right-3 opacity-50 hover:opacity-100" aria-label="Close"><X size={14} /></button>
              <p style={{ color: COLORS.graphite }} className="font-medium mb-1">Preview mode</p>
              <p className="leading-relaxed pr-4">Any email/password gets you in — this demonstrates the flow before real authentication is wired in.</p>
            </div>
          )}

          <form onSubmit={handleSubmit} className="space-y-4">
            <div>
              <label className="block text-sm font-medium mb-1.5">Email</label>
              <input type="email" required value={email} onChange={(e) => setEmail(e.target.value)} placeholder="you@company.com" style={{ borderColor: COLORS.line }} className="w-full rounded-lg border px-3.5 py-2.5 text-sm outline-none" />
            </div>
            <div>
              <label className="block text-sm font-medium mb-1.5">Password</label>
              <div className="relative">
                <input type={showPassword ? "text" : "password"} required value={password} onChange={(e) => setPassword(e.target.value)} placeholder="••••••••" style={{ borderColor: COLORS.line }} className="w-full rounded-lg border px-3.5 py-2.5 pr-10 text-sm outline-none" />
                <button type="button" onClick={() => setShowPassword(!showPassword)} className="absolute right-3 top-1/2 -translate-y-1/2 opacity-50 hover:opacity-90" aria-label="Toggle password">
                  {showPassword ? <EyeOff size={16} /> : <Eye size={16} />}
                </button>
              </div>
            </div>
            <button type="submit" style={{ background: COLORS.graphite }} className="w-full flex items-center justify-center gap-2 text-white rounded-lg py-2.5 text-sm font-medium mt-2 hover:opacity-90">
              Sign in <ArrowRight size={15} />
            </button>
          </form>
        </div>
      </div>
    </div>
  );
}

// ---------- Add Product Modal ----------
function AddProductModal({ onClose, onAdd }) {
  const [form, setForm] = useState({ name: "", category: "Tiles", qty: "", threshold: "", price: "", supplier: "" });
  const update = (key) => (e) => setForm({ ...form, [key]: e.target.value });

  const handleSubmit = (e) => {
    e.preventDefault();
    if (!form.name || !form.qty || !form.price) return;
    onAdd({
      id: Date.now(), name: form.name, category: form.category, qty: Number(form.qty),
      threshold: Number(form.threshold) || 10, price: Number(form.price), supplier: form.supplier || "—",
      dateAdded: new Date().toISOString().slice(0, 10),
    });
    onClose();
  };

  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 px-4">
      <div style={{ background: COLORS.surface }} className="w-full max-w-md rounded-2xl p-6 md:p-7 relative max-h-[90vh] overflow-y-auto">
        <button onClick={onClose} className="absolute top-5 right-5 opacity-50 hover:opacity-100" aria-label="Close"><X size={18} /></button>
        <h2 style={{ fontFamily: "'Sora', sans-serif" }} className="text-xl font-semibold mb-1">Add product</h2>
        <p style={{ color: COLORS.mist }} className="text-sm mb-6">Enter details for the new stock item.</p>

        <div style={{ borderColor: COLORS.line, background: COLORS.porcelain }} className="border border-dashed rounded-xl h-24 flex flex-col items-center justify-center mb-5 gap-1">
          <ImageIcon size={18} style={{ color: COLORS.mist }} />
          <p style={{ color: COLORS.mist }} className="text-xs">Image upload — added once backend is connected</p>
        </div>

        <form onSubmit={handleSubmit} className="space-y-3.5">
          <div>
            <label className="block text-sm font-medium mb-1.5">Product name</label>
            <input required value={form.name} onChange={update("name")} placeholder="e.g. Carrara Matte 600x600" style={{ borderColor: COLORS.line }} className="w-full rounded-lg border px-3.5 py-2.5 text-sm outline-none" />
          </div>
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="block text-sm font-medium mb-1.5">Category</label>
              <select value={form.category} onChange={update("category")} style={{ borderColor: COLORS.line }} className="w-full rounded-lg border px-3.5 py-2.5 text-sm outline-none bg-white">
                <option>Tiles</option><option>Adhesive</option>
              </select>
            </div>
            <div>
              <label className="block text-sm font-medium mb-1.5">Supplier</label>
              <input value={form.supplier} onChange={update("supplier")} placeholder="Supplier name" style={{ borderColor: COLORS.line }} className="w-full rounded-lg border px-3.5 py-2.5 text-sm outline-none" />
            </div>
          </div>
          <div className="grid grid-cols-3 gap-3">
            <div>
              <label className="block text-sm font-medium mb-1.5">Quantity</label>
              <input required type="number" min="0" value={form.qty} onChange={update("qty")} style={{ borderColor: COLORS.line }} className="w-full rounded-lg border px-3 py-2.5 text-sm outline-none" />
            </div>
            <div>
              <label className="block text-sm font-medium mb-1.5">Low-stock at</label>
              <input type="number" min="0" value={form.threshold} onChange={update("threshold")} placeholder="10" style={{ borderColor: COLORS.line }} className="w-full rounded-lg border px-3 py-2.5 text-sm outline-none" />
            </div>
            <div>
              <label className="block text-sm font-medium mb-1.5">Price (₹)</label>
              <input required type="number" min="0" step="0.01" value={form.price} onChange={update("price")} style={{ borderColor: COLORS.line }} className="w-full rounded-lg border px-3 py-2.5 text-sm outline-none" />
            </div>
          </div>
          <button type="submit" style={{ background: COLORS.graphite }} className="w-full flex items-center justify-center gap-2 text-white rounded-lg py-2.5 text-sm font-medium mt-2 hover:opacity-90">
            <Check size={15} /> Save product
          </button>
        </form>
      </div>
    </div>
  );
}

function DeleteConfirm({ product, onCancel, onConfirm }) {
  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 px-4">
      <div style={{ background: COLORS.surface }} className="w-full max-w-sm rounded-2xl p-6">
        <div style={{ background: COLORS.rustFaint }} className="w-11 h-11 rounded-full flex items-center justify-center mb-4">
          <Trash2 size={18} color={COLORS.rust} />
        </div>
        <h2 className="font-semibold text-lg mb-1.5">Delete "{product.name}"?</h2>
        <p style={{ color: COLORS.mist }} className="text-sm mb-6">This removes it from your inventory. This can't be undone.</p>
        <div className="flex gap-3">
          <button onClick={onCancel} style={{ borderColor: COLORS.line }} className="flex-1 border rounded-lg py-2.5 text-sm font-medium hover:bg-gray-50">Cancel</button>
          <button onClick={onConfirm} style={{ background: COLORS.rust }} className="flex-1 text-white rounded-lg py-2.5 text-sm font-medium hover:opacity-90">Delete</button>
        </div>
      </div>
    </div>
  );
}

// ---------- Import Modal ----------
function ImportModal({ onClose, onImport }) {
  const fileRef = useRef(null);
  const [status, setStatus] = useState(null); // {type: 'success'|'error', message}

  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "binary" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet);
        const parsed = parseImportedRows(rows);
        onImport(parsed);
        setStatus({ type: "success", message: `Imported ${parsed.length} product${parsed.length === 1 ? "" : "s"}.` });
      } catch (err) {
        setStatus({ type: "error", message: "Couldn't read that file. Make sure it matches the template structure." });
      }
    };
    reader.readAsBinaryString(file);
  };

  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 px-4">
      <div style={{ background: COLORS.surface }} className="w-full max-w-md rounded-2xl p-6 md:p-7 relative">
        <button onClick={onClose} className="absolute top-5 right-5 opacity-50 hover:opacity-100" aria-label="Close"><X size={18} /></button>
        <h2 style={{ fontFamily: "'Sora', sans-serif" }} className="text-xl font-semibold mb-1">Import from Excel</h2>
        <p style={{ color: COLORS.mist }} className="text-sm mb-5">Upload a .xlsx file matching the required columns.</p>

        <div style={{ background: COLORS.tealFaint, border: `1px solid ${COLORS.teal}22` }} className="rounded-lg p-4 mb-5 text-xs">
          <p style={{ color: COLORS.graphite }} className="font-medium mb-1.5">Required columns</p>
          <p style={{ color: COLORS.ink }} className="leading-relaxed">
            Name, Category, Quantity, Price, Supplier, Low Stock Threshold, Date Added
          </p>
          <button onClick={downloadSampleTemplate} style={{ color: COLORS.teal }} className="flex items-center gap-1.5 font-medium mt-3 hover:underline">
            <Download size={13} /> Download sample template
          </button>
        </div>

        <label
          style={{ borderColor: COLORS.line, background: COLORS.porcelain }}
          className="border border-dashed rounded-xl h-28 flex flex-col items-center justify-center gap-1.5 cursor-pointer"
        >
          <Upload size={18} style={{ color: COLORS.mist }} />
          <p style={{ color: COLORS.mist }} className="text-xs">Click to choose a .xlsx file</p>
          <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleFile} className="hidden" />
        </label>

        {status && (
          <p
            style={{ color: status.type === "success" ? COLORS.teal : COLORS.rust, background: status.type === "success" ? COLORS.tealFaint : COLORS.rustFaint }}
            className="text-sm text-center rounded-lg py-2.5 px-3 mt-4"
          >
            {status.message}
          </p>
        )}
      </div>
    </div>
  );
}

// ---------- Top Nav ----------
function TopNav({ activeTab, setActiveTab, onLogout }) {
  const tabs = [
    { key: "dashboard", label: "Dashboard", icon: LayoutDashboard },
    { key: "products", label: "Products", icon: Package },
  ];
  return (
    <header style={{ background: COLORS.graphite }} className="text-white sticky top-0 z-20">
      <div className="max-w-6xl mx-auto px-5 md:px-8 flex items-center justify-between h-16">
        <div className="flex items-center gap-8">
          <div className="flex items-center gap-2.5">
            <div style={{ background: COLORS.teal }} className="w-8 h-8 rounded-md flex items-center justify-center">
              <Grid3x3 size={16} color="#fff" />
            </div>
            <span style={{ fontFamily: "'Sora', sans-serif" }} className="font-semibold hidden sm:inline">GroutLine</span>
          </div>
          <nav className="flex items-center gap-1">
            {tabs.map(({ key, label, icon: Icon }) => (
              <button
                key={key}
                onClick={() => setActiveTab(key)}
                style={{ borderBottom: activeTab === key ? `2px solid ${COLORS.tealLight}` : "2px solid transparent" }}
                className="flex items-center gap-2 px-3 py-5 text-sm font-medium transition-colors"
              >
                <Icon size={15} color={activeTab === key ? "#fff" : "rgba(255,255,255,0.55)"} />
                <span style={{ color: activeTab === key ? "#fff" : "rgba(255,255,255,0.55)" }}>{label}</span>
              </button>
            ))}
          </nav>
        </div>
        <div className="flex items-center gap-4">
          <button style={{ color: "rgba(255,255,255,0.6)" }} className="relative hover:opacity-80" aria-label="Notifications">
            <Bell size={17} />
            <span style={{ background: COLORS.rust }} className="absolute -top-1 -right-1 w-2 h-2 rounded-full" />
          </button>
          <button onClick={onLogout} style={{ color: "rgba(255,255,255,0.6)" }} className="hover:opacity-80" aria-label="Log out">
            <LogOut size={17} />
          </button>
          <div style={{ background: COLORS.tealLight }} className="w-8 h-8 rounded-full flex items-center justify-center text-white text-xs font-medium">A</div>
        </div>
      </div>
    </header>
  );
}

// ---------- Dashboard Tab (overview only — full browsable list lives in Products) ----------
function DashboardTab({ products, onAdd, onDelete, onImport, goToProducts }) {
  const [showAdd, setShowAdd] = useState(false);
  const [showImport, setShowImport] = useState(false);
  const [deleteTarget, setDeleteTarget] = useState(null);

  const stats = useMemo(() => {
    const totalValue = products.reduce((s, p) => s + p.qty * p.price, 0);
    const lowStock = products.filter((p) => p.qty < p.threshold);
    const suppliers = new Set(products.map((p) => p.supplier)).size;
    const recent = [...products].sort((a, b) => new Date(b.dateAdded) - new Date(a.dateAdded)).slice(0, 5);
    return { totalValue, lowStock, suppliers, recent };
  }, [products]);

  return (
    <div className="max-w-5xl mx-auto px-5 md:px-8 py-7 space-y-5">
      {/* Header */}
      <div className="flex items-start justify-between">
        <div>
          <h1 style={{ fontFamily: "'Sora', sans-serif" }} className="text-2xl font-semibold">Welcome back</h1>
          <p style={{ color: COLORS.mist }} className="text-sm mt-1">Quick snapshot — full list and filters live in Products.</p>
        </div>
        <button onClick={() => setShowAdd(true)} style={{ background: COLORS.graphite }} className="hidden sm:flex items-center gap-2 text-white rounded-lg px-4 py-2.5 text-sm font-medium hover:opacity-90 flex-shrink-0">
          <Plus size={15} /> Add product
        </button>
      </div>

      {/* Stats */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
        <StatCard icon={Boxes} label="Total products" value={products.length} accent={{ solid: COLORS.teal, faint: COLORS.tealFaint }} />
        <StatCard icon={DollarSign} label="Inventory value" value={formatINR(stats.totalValue)} accent={{ solid: COLORS.sand, faint: COLORS.sandFaint }} />
        <StatCard icon={AlertTriangle} label="Low stock" value={stats.lowStock.length} accent={{ solid: COLORS.rust, faint: COLORS.rustFaint }} />
        <StatCard icon={Users} label="Suppliers" value={stats.suppliers} accent={{ solid: COLORS.graphite, faint: "#EFEFEE" }} />
      </div>

      {/* Low stock (compact, scrolls internally if long) + Excel actions */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <div style={{ background: COLORS.surface, border: `1px solid ${COLORS.line}` }} className="rounded-xl p-5">
          <div className="flex items-center justify-between mb-3">
            <div className="flex items-center gap-2">
              <AlertTriangle size={14} color={COLORS.rust} />
              <h2 className="font-semibold text-sm">Needs attention</h2>
            </div>
            {stats.lowStock.length > 0 && (
              <button onClick={goToProducts} style={{ color: COLORS.teal }} className="text-xs font-medium hover:underline">View all</button>
            )}
          </div>
          {stats.lowStock.length === 0 ? (
            <p style={{ color: COLORS.mist }} className="text-sm">Nothing below threshold right now.</p>
          ) : (
            <div className="space-y-2.5 max-h-40 overflow-y-auto pr-1">
              {stats.lowStock.map((p) => (
                <div key={p.id} className="flex items-center justify-between">
                  <p className="text-sm truncate pr-3">{p.name}</p>
                  <span style={{ background: COLORS.rustFaint, color: COLORS.rust, fontFamily: "'IBM Plex Mono', monospace" }} className="text-xs px-2 py-0.5 rounded-full flex-shrink-0">
                    {p.qty}/{p.threshold}
                  </span>
                </div>
              ))}
            </div>
          )}
        </div>

        <div style={{ background: COLORS.surface, border: `1px solid ${COLORS.line}` }} className="rounded-xl p-5 flex flex-col gap-2.5">
          <h2 className="font-semibold text-sm mb-0.5">Excel</h2>
          <button onClick={() => exportToExcel(products, "groutline-inventory-export.xlsx")} style={{ background: COLORS.teal }} className="w-full flex items-center justify-center gap-2 text-white rounded-lg py-2.5 text-sm font-medium hover:opacity-90">
            <Download size={15} /> Export all to Excel
          </button>
          <button onClick={() => setShowImport(true)} style={{ borderColor: COLORS.line }} className="w-full flex items-center justify-center gap-2 border rounded-lg py-2.5 text-sm font-medium hover:bg-gray-50">
            <Upload size={15} /> Import from Excel
          </button>
        </div>
      </div>

      {/* Recently added — short list, not the full table */}
      <div style={{ background: COLORS.surface, border: `1px solid ${COLORS.line}` }} className="rounded-xl overflow-hidden">
        <div style={{ borderBottom: `1px solid ${COLORS.line}` }} className="flex items-center justify-between px-5 py-3.5">
          <h2 className="font-semibold text-sm">Recently added</h2>
          <button onClick={goToProducts} style={{ color: COLORS.teal }} className="text-xs font-medium hover:underline">View all products →</button>
        </div>
        <div>
          {stats.recent.map((p, i) => {
            const low = p.qty < p.threshold;
            return (
              <div key={p.id} style={{ borderTop: i === 0 ? "none" : `1px solid ${COLORS.line}` }} className="flex items-center justify-between px-5 py-3">
                <div className="flex items-center gap-3 min-w-0">
                  <ProductVisual category={p.category} size={30} />
                  <div className="min-w-0">
                    <p className="text-sm font-medium truncate">{p.name}</p>
                    <p style={{ color: COLORS.mist }} className="text-xs">{p.dateAdded}</p>
                  </div>
                </div>
                <span style={{ color: low ? COLORS.rust : COLORS.mist, fontFamily: "'IBM Plex Mono', monospace" }} className="text-xs flex-shrink-0">{p.qty} in stock</span>
              </div>
            );
          })}
        </div>
      </div>

      <button onClick={() => setShowAdd(true)} className="sm:hidden w-full flex items-center justify-center gap-2 rounded-lg py-2.5 text-sm font-medium" style={{ background: COLORS.graphite, color: "#fff" }}>
        <Plus size={15} /> Add product
      </button>

      {showAdd && <AddProductModal onClose={() => setShowAdd(false)} onAdd={onAdd} />}
      {showImport && <ImportModal onClose={() => setShowImport(false)} onImport={onImport} />}
      {deleteTarget && <DeleteConfirm product={deleteTarget} onCancel={() => setDeleteTarget(null)} onConfirm={() => { onDelete(deleteTarget.id); setDeleteTarget(null); }} />}
    </div>
  );
}

// ---------- Products Tab ----------
function ProductsTab({ products, onAdd, onDelete, onImport }) {
  const [search, setSearch] = useState("");
  const [category, setCategory] = useState("All");
  const [lowStockOnly, setLowStockOnly] = useState(false);
  const [sortBy, setSortBy] = useState("newest");
  const [showAdd, setShowAdd] = useState(false);
  const [showImport, setShowImport] = useState(false);
  const [deleteTarget, setDeleteTarget] = useState(null);

  const filtered = useMemo(() => {
    let list = products.filter((p) => {
      const matchesSearch = p.name.toLowerCase().includes(search.toLowerCase()) || p.supplier.toLowerCase().includes(search.toLowerCase());
      const matchesCategory = category === "All" || p.category === category;
      const matchesLowStock = !lowStockOnly || p.qty < p.threshold;
      return matchesSearch && matchesCategory && matchesLowStock;
    });
    list = [...list].sort((a, b) => {
      if (sortBy === "newest") return new Date(b.dateAdded) - new Date(a.dateAdded);
      if (sortBy === "oldest") return new Date(a.dateAdded) - new Date(b.dateAdded);
      if (sortBy === "name") return a.name.localeCompare(b.name);
      if (sortBy === "qty") return a.qty - b.qty;
      return 0;
    });
    return list;
  }, [products, search, category, lowStockOnly, sortBy]);

  const filtersActive = search || category !== "All" || lowStockOnly;

  return (
    <div className="max-w-6xl mx-auto px-5 md:px-8 py-7 space-y-5">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
        <div>
          <h1 style={{ fontFamily: "'Sora', sans-serif" }} className="text-2xl font-semibold">Products</h1>
          <p style={{ color: COLORS.mist }} className="text-sm mt-1">{filtered.length} of {products.length} items shown</p>
        </div>
        <div className="flex gap-2">
          <button onClick={() => setShowImport(true)} style={{ borderColor: COLORS.line }} className="flex items-center gap-2 border rounded-lg px-3.5 py-2.5 text-sm font-medium hover:bg-gray-50">
            <Upload size={14} /> Import
          </button>
          <button onClick={() => setShowAdd(true)} style={{ background: COLORS.graphite }} className="flex items-center gap-2 text-white rounded-lg px-3.5 py-2.5 text-sm font-medium hover:opacity-90">
            <Plus size={14} /> Add
          </button>
        </div>
      </div>

      {/* Filters */}
      <div style={{ background: COLORS.surface, border: `1px solid ${COLORS.line}` }} className="rounded-xl p-4 flex flex-wrap items-center gap-3">
        <div className="relative flex-1 min-w-[160px]">
          <Search size={14} style={{ color: COLORS.mist }} className="absolute left-3 top-1/2 -translate-y-1/2" />
          <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search name or supplier..." style={{ borderColor: COLORS.line }} className="w-full pl-8 pr-3 py-2 rounded-lg border text-sm outline-none" />
        </div>

        <select value={category} onChange={(e) => setCategory(e.target.value)} style={{ borderColor: COLORS.line }} className="rounded-lg border px-3 py-2 text-sm outline-none bg-white">
          <option value="All">All categories</option>
          <option value="Tiles">Tiles</option>
          <option value="Adhesive">Adhesive</option>
        </select>

        <select value={sortBy} onChange={(e) => setSortBy(e.target.value)} style={{ borderColor: COLORS.line }} className="rounded-lg border px-3 py-2 text-sm outline-none bg-white">
          <option value="newest">Newest first</option>
          <option value="oldest">Oldest first</option>
          <option value="name">Name A–Z</option>
          <option value="qty">Quantity, low to high</option>
        </select>

        <label className="flex items-center gap-2 text-sm cursor-pointer select-none px-1">
          <input type="checkbox" checked={lowStockOnly} onChange={(e) => setLowStockOnly(e.target.checked)} style={{ accentColor: COLORS.rust }} className="w-4 h-4 rounded" />
          Low stock only
        </label>

        <button
          onClick={() => exportToExcel(filtered, filtersActive ? "groutline-filtered-export.xlsx" : "groutline-inventory-export.xlsx")}
          style={{ background: COLORS.teal }}
          className="ml-auto flex items-center gap-2 text-white rounded-lg px-3.5 py-2 text-sm font-medium hover:opacity-90"
        >
          <FileSpreadsheet size={14} /> Export {filtersActive ? "filtered" : "all"}
        </button>
      </div>

      {filtered.length === 0 ? (
        <div style={{ background: COLORS.surface, border: `1px solid ${COLORS.line}` }} className="rounded-xl py-16 text-center">
          <Boxes size={28} style={{ color: COLORS.mist }} className="mx-auto mb-3" />
          <p style={{ color: COLORS.mist }} className="text-sm">No products match these filters.</p>
        </div>
      ) : (
        <div style={{ background: COLORS.surface, border: `1px solid ${COLORS.line}` }} className="rounded-xl overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr style={{ color: COLORS.mist, borderBottom: `1px solid ${COLORS.line}` }} className="text-left">
                  <th className="font-medium px-5 py-3">Product</th>
                  <th className="font-medium px-5 py-3">Qty</th>
                  <th className="font-medium px-5 py-3">Price</th>
                  <th className="font-medium px-5 py-3">Supplier</th>
                  <th className="font-medium px-5 py-3">Added</th>
                  <th className="font-medium px-5 py-3"></th>
                </tr>
              </thead>
              <tbody>
                {filtered.map((p) => {
                  const low = p.qty < p.threshold;
                  return (
                    <tr key={p.id} style={{ borderTop: `1px solid ${COLORS.line}` }}>
                      <td className="px-5 py-3">
                        <div className="flex items-center gap-3">
                          <ProductVisual category={p.category} size={34} />
                          <div>
                            <p className="font-medium">{p.name}</p>
                            <p style={{ color: COLORS.mist }} className="text-xs">{p.category}</p>
                          </div>
                        </div>
                      </td>
                      <td className="px-5 py-3">
                        <span style={{ color: low ? COLORS.rust : COLORS.ink, fontFamily: "'IBM Plex Mono', monospace" }} className="text-xs">{p.qty}</span>
                        {low && <span style={{ color: COLORS.rust }} className="text-xs ml-1.5">low</span>}
                      </td>
                      <td style={{ fontFamily: "'IBM Plex Mono', monospace" }} className="px-5 py-3 text-xs">{formatINR(p.price)}</td>
                      <td style={{ color: COLORS.mist }} className="px-5 py-3 text-xs">{p.supplier}</td>
                      <td style={{ color: COLORS.mist, fontFamily: "'IBM Plex Mono', monospace" }} className="px-5 py-3 text-xs">{p.dateAdded}</td>
                      <td className="px-5 py-3 text-right">
                        <button onClick={() => setDeleteTarget(p)} className="opacity-40 hover:opacity-100 hover:text-red-600 transition-opacity" aria-label={`Delete ${p.name}`}>
                          <Trash2 size={15} />
                        </button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {showAdd && <AddProductModal onClose={() => setShowAdd(false)} onAdd={onAdd} />}
      {showImport && <ImportModal onClose={() => setShowImport(false)} onImport={onImport} />}
      {deleteTarget && <DeleteConfirm product={deleteTarget} onCancel={() => setDeleteTarget(null)} onConfirm={() => { onDelete(deleteTarget.id); setDeleteTarget(null); }} />}
    </div>
  );
}

// ---------- Root App ----------
export default function App() {
  const [loggedIn, setLoggedIn] = useState(false);
  const [activeTab, setActiveTab] = useState("dashboard");
  const [products, setProducts] = useState(INITIAL_PRODUCTS);

  const handleAdd = (product) => setProducts((prev) => [product, ...prev]);
  const handleDelete = (id) => setProducts((prev) => prev.filter((p) => p.id !== id));
  const handleImport = (newProducts) => setProducts((prev) => [...newProducts, ...prev]);

  if (!loggedIn) return <LoginScreen onLogin={() => setLoggedIn(true)} />;

  return (
    <div
      style={{ fontFamily: "'Inter', sans-serif", color: COLORS.ink, background: COLORS.porcelain, height: "100vh", overflowY: "auto" }}
    >
      <style>{FONT_IMPORT}</style>
      <TopNav activeTab={activeTab} setActiveTab={(t) => setActiveTab(t)} onLogout={() => setLoggedIn(false)} />
      {activeTab === "dashboard" ? (
        <DashboardTab products={products} onAdd={handleAdd} onDelete={handleDelete} onImport={handleImport} goToProducts={() => setActiveTab("products")} />
      ) : (
        <ProductsTab products={products} onAdd={handleAdd} onDelete={handleDelete} onImport={handleImport} />
      )}
    </div>
  );
}
