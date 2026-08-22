import React, { useState, useMemo, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  Eye, EyeOff, Info, Grid3x3, ArrowRight, X, LayoutDashboard, Package,
  Plus, Trash2, Search, Bell, AlertTriangle, DollarSign, Users, Boxes,
  Layers, Droplet, Download, Upload, FileSpreadsheet, LogOut, Check,
  ChevronDown, Image as ImageIcon, Pencil, ExternalLink, RefreshCw, AlertCircle,
  UserPlus, Mail, Lock,
} from "lucide-react";
import { GOOGLE_CLIENT_ID } from "./config";
import { initGoogleAuth, signIn, signOut, trySilentSignIn } from "./services/googleAuthService";
import { getOrCreateUserSheet, readProducts, writeAllProducts, sheetUrl } from "./services/sheetsService";
import { backupProducts } from "./services/backupService";
import { inviteTeamMember, signInWithEmail } from "./services/teamAuthService";
import { readProductsViaBridge, writeProductsViaBridge } from "./services/sheetBridgeService";

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

function GoogleGlyph() {
  return (
    <svg width="16" height="16" viewBox="0 0 48 48" aria-hidden="true">
      <path fill="#FFC107" d="M43.6 20.5H42V20H24v8h11.3C33.7 32.9 29.3 36 24 36c-6.6 0-12-5.4-12-12s5.4-12 12-12c3.1 0 5.8 1.1 8 3l6-6C33.9 5.5 29.2 3.5 24 3.5 12.7 3.5 3.5 12.7 3.5 24S12.7 44.5 24 44.5 44.5 35.3 44.5 24c0-1.2-.1-2.4-.4-3.5z"/>
      <path fill="#FF3D00" d="M6.3 14.7l6.6 4.8C14.6 15.9 18.9 13 24 13c3.1 0 5.8 1.1 8 3l6-6C33.9 5.5 29.2 3.5 24 3.5c-7.7 0-14.4 4.3-17.7 11.2z"/>
      <path fill="#4CAF50" d="M24 44.5c5.1 0 9.8-1.9 13.3-5.1l-6.2-5.2c-2 1.4-4.6 2.3-7.1 2.3-5.3 0-9.7-3.1-11.3-7.8l-6.5 5C9.5 40.1 16.2 44.5 24 44.5z"/>
      <path fill="#1976D2" d="M43.6 20.5H42V20H24v8h11.3c-.8 2.4-2.3 4.4-4.2 5.9l6.2 5.2C40.9 36 44.5 30.5 44.5 24c0-1.2-.1-2.4-.4-3.5z"/>
    </svg>
  );
}

function LoginScreen({ onSignedIn, googleReady }) {
  const [mode, setMode] = useState("google"); // "google" | "team"
  const [showInfo, setShowInfo] = useState(false);
  const [status, setStatus] = useState("idle"); // idle | signing-in | loading-sheet | error
  const [errorMsg, setErrorMsg] = useState("");
  const [teamEmail, setTeamEmail] = useState("");
  const [teamPassword, setTeamPassword] = useState("");
  const [showPassword, setShowPassword] = useState(false);

  const handleGoogleSignIn = async () => {
    setStatus("signing-in");
    setErrorMsg("");
    try {
      const { token, user } = await signIn();
      setStatus("loading-sheet");
      const { sheetId, isNew } = await getOrCreateUserSheet(token);
      let products = [];
      if (isNew) {
        await writeAllProducts(token, sheetId, []);
      } else {
        products = await readProducts(token, sheetId);
      }
      onSignedIn({ mode: "google", token, user, sheetId, products });
    } catch (err) {
      console.error(err);
      setStatus("error");
      setErrorMsg(
        err?.message?.includes("popup") || err?.error === "popup_closed_by_user"
          ? "Sign-in was closed before finishing. Try again."
          : "Couldn't connect to Google. Check that this domain is added as an authorized origin in Google Cloud Console."
      );
    }
  };

  const handleTeamSignIn = async (e) => {
    e.preventDefault();
    setStatus("signing-in");
    setErrorMsg("");
    try {
      const { supabaseToken, email } = await signInWithEmail(teamEmail, teamPassword);
      setStatus("loading-sheet");
      const products = await readProductsViaBridge(supabaseToken);
      onSignedIn({ mode: "team", supabaseToken, user: { email }, products });
    } catch (err) {
      setStatus("error");
      setErrorMsg(err.message || "Sign-in failed.");
    }
  };

  const busy = status === "signing-in" || status === "loading-sheet";

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
          <p style={{ color: COLORS.mist }} className="text-sm mb-5">
            {mode === "google" ? "Your inventory lives in a Google Sheet in your own Drive." : "Sign in with the login your team owner shared with you."}
          </p>

          {showInfo && (
            <div style={{ background: "#EFF6F6", border: `1px solid ${COLORS.teal}22` }} className="rounded-lg p-4 mb-5 text-sm relative">
              <button onClick={() => setShowInfo(false)} className="absolute top-3 right-3 opacity-50 hover:opacity-100" aria-label="Close"><X size={14} /></button>
              <p style={{ color: COLORS.graphite }} className="font-medium mb-1">How this works</p>
              <p className="leading-relaxed pr-4">
                {mode === "google"
                  ? "Signing in creates (or opens) a Google Sheet called \"GroutLine Inventory\" in your Drive. Everything you do here writes straight to that sheet — you can also open and edit it directly in Google Sheets."
                  : "Team logins don't need a Google account — the owner creates this email and password for you from inside the app, and it gives you the same read/write access to the same inventory."}
              </p>
            </div>
          )}

          {/* Mode toggle */}
          <div style={{ border: `1px solid ${COLORS.line}`, background: COLORS.surface }} className="flex rounded-lg p-1 mb-5">
            <button
              onClick={() => { setMode("google"); setStatus("idle"); }}
              style={{ background: mode === "google" ? COLORS.graphite : "transparent", color: mode === "google" ? "#fff" : COLORS.mist }}
              className="flex-1 rounded-md py-2 text-xs font-medium transition-colors"
            >
              Google account
            </button>
            <button
              onClick={() => { setMode("team"); setStatus("idle"); }}
              style={{ background: mode === "team" ? COLORS.graphite : "transparent", color: mode === "team" ? "#fff" : COLORS.mist }}
              className="flex-1 rounded-md py-2 text-xs font-medium transition-colors"
            >
              Email & password
            </button>
          </div>

          {mode === "google" ? (
            <button
              onClick={handleGoogleSignIn}
              disabled={!googleReady || busy}
              style={{ borderColor: COLORS.line }}
              className="w-full flex items-center justify-center gap-3 border rounded-lg py-2.5 text-sm font-medium hover:bg-gray-50 disabled:opacity-60 disabled:cursor-not-allowed"
            >
              {busy ? (
                <>
                  <RefreshCw size={16} className="animate-spin" style={{ color: COLORS.teal }} />
                  {status === "signing-in" ? "Waiting for Google..." : "Loading your inventory..."}
                </>
              ) : (
                <>
                  <GoogleGlyph />
                  Sign in with Google
                </>
              )}
            </button>
          ) : (
            <form onSubmit={handleTeamSignIn} className="space-y-3">
              <div className="relative">
                <Mail size={15} style={{ color: COLORS.mist }} className="absolute left-3 top-1/2 -translate-y-1/2" />
                <input
                  type="email" required value={teamEmail} onChange={(e) => setTeamEmail(e.target.value)}
                  placeholder="teammate@company.com" style={{ borderColor: COLORS.line }}
                  className="w-full pl-9 pr-3 py-2.5 rounded-lg border text-sm outline-none"
                />
              </div>
              <div className="relative">
                <Lock size={15} style={{ color: COLORS.mist }} className="absolute left-3 top-1/2 -translate-y-1/2" />
                <input
                  type={showPassword ? "text" : "password"} required value={teamPassword} onChange={(e) => setTeamPassword(e.target.value)}
                  placeholder="Password" style={{ borderColor: COLORS.line }}
                  className="w-full pl-9 pr-10 py-2.5 rounded-lg border text-sm outline-none"
                />
                <button type="button" onClick={() => setShowPassword(!showPassword)} className="absolute right-3 top-1/2 -translate-y-1/2 opacity-50 hover:opacity-90" aria-label="Toggle password">
                  {showPassword ? <EyeOff size={15} /> : <Eye size={15} />}
                </button>
              </div>
              <button
                type="submit"
                disabled={busy}
                style={{ background: COLORS.graphite }}
                className="w-full flex items-center justify-center gap-2 text-white rounded-lg py-2.5 text-sm font-medium hover:opacity-90 disabled:opacity-60"
              >
                {busy ? <RefreshCw size={15} className="animate-spin" /> : <ArrowRight size={15} />}
                {busy ? (status === "signing-in" ? "Signing in..." : "Loading inventory...") : "Sign in"}
              </button>
            </form>
          )}

          {status === "error" && (
            <div style={{ background: COLORS.rustFaint, color: COLORS.rust }} className="flex items-start gap-2 rounded-lg p-3.5 mt-4 text-sm">
              <AlertCircle size={16} className="flex-shrink-0 mt-0.5" />
              <p>{errorMsg}</p>
            </div>
          )}

          {mode === "google" && !googleReady && (
            <p style={{ color: COLORS.mist }} className="text-xs text-center mt-4">Connecting to Google...</p>
          )}

          <div className="flex items-center justify-center gap-3 mt-8">
            <a href="/privacy.html" target="_blank" rel="noreferrer" style={{ color: COLORS.mist }} className="text-xs hover:underline">Privacy Policy</a>
            <span style={{ color: COLORS.line }}>·</span>
            <a href="/terms.html" target="_blank" rel="noreferrer" style={{ color: COLORS.mist }} className="text-xs hover:underline">Terms of Service</a>
          </div>
        </div>
      </div>
    </div>
  );
}

// ---------- Add Product Modal ----------
function ProductFormModal({ onClose, onSave, editProduct }) {
  const isEdit = Boolean(editProduct);
  const [form, setForm] = useState(
    editProduct
      ? {
          name: editProduct.name, category: editProduct.category, qty: String(editProduct.qty),
          threshold: String(editProduct.threshold), price: String(editProduct.price), supplier: editProduct.supplier,
        }
      : { name: "", category: "Tiles", qty: "", threshold: "", price: "", supplier: "" }
  );
  const update = (key) => (e) => setForm({ ...form, [key]: e.target.value });

  const handleSubmit = (e) => {
    e.preventDefault();
    if (!form.name || !form.qty || !form.price) return;
    onSave({
      id: isEdit ? editProduct.id : Date.now(),
      name: form.name,
      category: form.category,
      qty: Number(form.qty),
      threshold: Number(form.threshold) || 10,
      price: Number(form.price),
      supplier: form.supplier || "—",
      dateAdded: isEdit ? editProduct.dateAdded : new Date().toISOString().slice(0, 10),
    });
    onClose();
  };

  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 px-4" onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: COLORS.surface }} className="w-full max-w-md rounded-2xl p-6 md:p-7 relative max-h-[90vh] overflow-y-auto">
        <button onClick={onClose} className="absolute top-5 right-5 opacity-50 hover:opacity-100" aria-label="Close"><X size={18} /></button>
        <h2 style={{ fontFamily: "'Sora', sans-serif" }} className="text-xl font-semibold mb-1">{isEdit ? "Edit product" : "Add product"}</h2>
        <p style={{ color: COLORS.mist }} className="text-sm mb-6">{isEdit ? "Update details for this stock item." : "Enter details for the new stock item."}</p>

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
            <Check size={15} /> {isEdit ? "Save changes" : "Save product"}
          </button>
        </form>
      </div>
    </div>
  );
}

function EditableQty({ product, onSave }) {
  const [editing, setEditing] = useState(false);
  const [value, setValue] = useState(product.qty);
  const low = product.qty < product.threshold;

  const commit = () => {
    const n = Number(value);
    if (!Number.isNaN(n) && n >= 0) onSave({ ...product, qty: n });
    setEditing(false);
  };

  if (editing) {
    return (
      <input
        type="number"
        min="0"
        autoFocus
        value={value}
        onChange={(e) => setValue(e.target.value)}
        onBlur={commit}
        onKeyDown={(e) => {
          if (e.key === "Enter") commit();
          if (e.key === "Escape") { setValue(product.qty); setEditing(false); }
        }}
        style={{ borderColor: COLORS.teal, fontFamily: "'IBM Plex Mono', monospace" }}
        className="w-16 border rounded-md px-2 py-1 text-xs outline-none"
      />
    );
  }

  return (
    <button
      onClick={() => setEditing(true)}
      style={{ color: low ? COLORS.rust : COLORS.ink, fontFamily: "'IBM Plex Mono', monospace" }}
      className="text-xs underline decoration-dotted underline-offset-2 hover:opacity-70"
      title="Click to edit quantity"
    >
      {product.qty}
    </button>
  );
}

function DeleteConfirm({ product, onCancel, onConfirm }) {
  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 px-4" onClick={onCancel}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: COLORS.surface }} className="w-full max-w-sm rounded-2xl p-6">
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
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 px-4" onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: COLORS.surface }} className="w-full max-w-md rounded-2xl p-6 md:p-7 relative">
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

// ---------- Share Access Modal (owner only) ----------
function ShareAccessModal({ onClose, ownerEmail, sheetId }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [status, setStatus] = useState("idle"); // idle | saving | success | error
  const [errorMsg, setErrorMsg] = useState("");

  const handleSubmit = async (e) => {
    e.preventDefault();
    setStatus("saving");
    setErrorMsg("");
    try {
      await inviteTeamMember({ ownerEmail, sheetId, teammateEmail: email, password });
      setStatus("success");
    } catch (err) {
      setStatus("error");
      setErrorMsg(err.message);
    }
  };

  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 px-4" style={{ color: COLORS.ink }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: COLORS.surface }} className="w-full max-w-md rounded-2xl p-6 md:p-7 relative">
        <button onClick={onClose} className="absolute top-5 right-5 opacity-50 hover:opacity-100" aria-label="Close"><X size={18} /></button>
        <h2 style={{ fontFamily: "'Sora', sans-serif" }} className="text-xl font-semibold mb-1">Share access</h2>
        <p style={{ color: COLORS.mist }} className="text-sm mb-6">
          Create a login for a teammate — they'll see and update the same inventory without needing a Google account.
        </p>

        {status === "success" ? (
          <div>
            <div style={{ background: COLORS.tealFaint }} className="rounded-lg p-4 text-sm mb-5">
              <p style={{ color: COLORS.graphite }} className="font-medium mb-1">Access created</p>
              <p style={{ color: COLORS.ink }}>Share these credentials with them directly — email: <span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{email}</span></p>
            </div>
            <button onClick={onClose} style={{ background: COLORS.graphite }} className="w-full text-white rounded-lg py-2.5 text-sm font-medium hover:opacity-90">
              Done
            </button>
          </div>
        ) : (
          <form onSubmit={handleSubmit} className="space-y-3.5">
            <div>
              <label className="block text-sm font-medium mb-1.5">Teammate's email</label>
              <input required type="email" value={email} onChange={(e) => setEmail(e.target.value)} placeholder="teammate@company.com" style={{ borderColor: COLORS.line }} className="w-full rounded-lg border px-3.5 py-2.5 text-sm outline-none" />
            </div>
            <div>
              <label className="block text-sm font-medium mb-1.5">Set a password for them</label>
              <input required type="text" minLength={8} value={password} onChange={(e) => setPassword(e.target.value)} placeholder="At least 8 characters" style={{ borderColor: COLORS.line }} className="w-full rounded-lg border px-3.5 py-2.5 text-sm outline-none" />
            </div>
            {status === "error" && (
              <p style={{ background: COLORS.rustFaint, color: COLORS.rust }} className="text-sm rounded-lg p-3">{errorMsg}</p>
            )}
            <button type="submit" disabled={status === "saving"} style={{ background: COLORS.graphite }} className="w-full flex items-center justify-center gap-2 text-white rounded-lg py-2.5 text-sm font-medium mt-2 hover:opacity-90 disabled:opacity-60">
              {status === "saving" ? <RefreshCw size={15} className="animate-spin" /> : <UserPlus size={15} />}
              {status === "saving" ? "Creating access..." : "Create access"}
            </button>
          </form>
        )}
      </div>
    </div>
  );
}

// ---------- Top Nav ----------
function TopNav({ activeTab, setActiveTab, onLogout, user, sheetId, syncStatus, sessionMode }) {
  const [showShare, setShowShare] = useState(false);
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
          {/* Sync status */}
          <div className="hidden md:flex items-center gap-1.5" style={{ color: "rgba(255,255,255,0.5)" }}>
            {syncStatus === "saving" && (
              <>
                <RefreshCw size={13} className="animate-spin" />
                <span className="text-xs">Saving...</span>
              </>
            )}
            {syncStatus === "saved" && (
              <>
                <Check size={13} color={COLORS.tealLight} />
                <span className="text-xs">Synced to Sheets</span>
              </>
            )}
            {syncStatus === "error" && (
              <>
                <AlertCircle size={13} color="#E8998D" />
                <span className="text-xs">Sync failed</span>
              </>
            )}
          </div>

          {sheetId && (
            <a
              href={sheetUrl(sheetId)}
              target="_blank"
              rel="noreferrer"
              style={{ color: "rgba(255,255,255,0.6)" }}
              className="hidden sm:flex items-center gap-1.5 hover:opacity-80 text-xs"
              title="Open in Google Sheets"
            >
              <ExternalLink size={14} />
              <span className="hidden lg:inline">Open in Sheets</span>
            </a>
          )}

          {sessionMode === "google" && (
            <button
              onClick={() => setShowShare(true)}
              style={{ color: "rgba(255,255,255,0.6)" }}
              className="hidden sm:flex items-center gap-1.5 hover:opacity-80"
              title="Share access with a teammate"
            >
              <UserPlus size={16} />
            </button>
          )}

          <button style={{ color: "rgba(255,255,255,0.6)" }} className="relative hover:opacity-80" aria-label="Notifications">
            <Bell size={17} />
            <span style={{ background: COLORS.rust }} className="absolute -top-1 -right-1 w-2 h-2 rounded-full" />
          </button>
          <button onClick={onLogout} style={{ color: "rgba(255,255,255,0.6)" }} className="hover:opacity-80" aria-label="Log out">
            <LogOut size={17} />
          </button>
          {user?.picture ? (
            <img src={user.picture} alt={user.name || "User"} title={user.email} className="w-8 h-8 rounded-full" referrerPolicy="no-referrer" />
          ) : (
            <div style={{ background: COLORS.tealLight }} className="w-8 h-8 rounded-full flex items-center justify-center text-white text-xs font-medium">
              {user?.email?.[0]?.toUpperCase() || "A"}
            </div>
          )}
        </div>
      </div>
      {showShare && <ShareAccessModal onClose={() => setShowShare(false)} ownerEmail={user?.email} sheetId={sheetId} />}
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

      {showAdd && <ProductFormModal onClose={() => setShowAdd(false)} onSave={onAdd} />}
      {showImport && <ImportModal onClose={() => setShowImport(false)} onImport={onImport} />}
      {deleteTarget && <DeleteConfirm product={deleteTarget} onCancel={() => setDeleteTarget(null)} onConfirm={() => { onDelete(deleteTarget.id); setDeleteTarget(null); }} />}
    </div>
  );
}

// ---------- Products Tab ----------
function ProductsTab({ products, onAdd, onEdit, onDelete, onImport }) {
  const [search, setSearch] = useState("");
  const [category, setCategory] = useState("All");
  const [lowStockOnly, setLowStockOnly] = useState(false);
  const [sortBy, setSortBy] = useState("newest");
  const [showAdd, setShowAdd] = useState(false);
  const [showImport, setShowImport] = useState(false);
  const [editTarget, setEditTarget] = useState(null);
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
                  <th className="font-medium px-5 py-3">Total value</th>
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
                        <EditableQty product={p} onSave={onEdit} />
                        {low && <span style={{ color: COLORS.rust }} className="text-xs ml-1.5">low</span>}
                      </td>
                      <td style={{ fontFamily: "'IBM Plex Mono', monospace" }} className="px-5 py-3 text-xs">{formatINR(p.price)}</td>
                      <td style={{ fontFamily: "'IBM Plex Mono', monospace", color: COLORS.teal }} className="px-5 py-3 text-xs font-medium">{formatINR(p.qty * p.price)}</td>
                      <td style={{ color: COLORS.mist }} className="px-5 py-3 text-xs">{p.supplier}</td>
                      <td style={{ color: COLORS.mist, fontFamily: "'IBM Plex Mono', monospace" }} className="px-5 py-3 text-xs">{p.dateAdded}</td>
                      <td className="px-5 py-3 text-right">
                        <div className="flex items-center justify-end gap-3">
                          <button onClick={() => setEditTarget(p)} style={{ color: COLORS.teal }} className="opacity-60 hover:opacity-100 transition-opacity" aria-label={`Edit ${p.name}`}>
                            <Pencil size={14} />
                          </button>
                          <button onClick={() => setDeleteTarget(p)} className="opacity-40 hover:opacity-100 hover:text-red-600 transition-opacity" aria-label={`Delete ${p.name}`}>
                            <Trash2 size={15} />
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {showAdd && <ProductFormModal onClose={() => setShowAdd(false)} onSave={onAdd} />}
      {editTarget && <ProductFormModal editProduct={editTarget} onClose={() => setEditTarget(null)} onSave={onEdit} />}
      {showImport && <ImportModal onClose={() => setShowImport(false)} onImport={onImport} />}
      {deleteTarget && <DeleteConfirm product={deleteTarget} onCancel={() => setDeleteTarget(null)} onConfirm={() => { onDelete(deleteTarget.id); setDeleteTarget(null); }} />}
    </div>
  );
}

// ---------- Root App ----------
const SESSION_KEY = "groutline_session_mode"; // "google" | "team"
const TEAM_TOKEN_KEY = "groutline_team_token";
const TEAM_EMAIL_KEY = "groutline_team_email";

export default function App() {
  const [loggedIn, setLoggedIn] = useState(false);
  const [activeTab, setActiveTab] = useState("dashboard");
  const [products, setProducts] = useState([]);

  const [googleReady, setGoogleReady] = useState(false);
  const [checkingSession, setCheckingSession] = useState(true);
  const [session, setSession] = useState(null); // { token, user, sheetId }
  const [syncStatus, setSyncStatus] = useState("idle"); // idle | saving | saved | error

  // Persist to localStorage on sign-in, restore on page load, clear on logout.
  const handleSignedIn = (payload, { persist = true } = {}) => {
    setSession(payload); // { mode: 'google', token, user, sheetId } OR { mode: 'team', supabaseToken, user }
    setProducts(payload.products);
    setLoggedIn(true);
    if (persist) {
      if (payload.mode === "google") {
        localStorage.setItem(SESSION_KEY, "google");
        localStorage.removeItem(TEAM_TOKEN_KEY);
        localStorage.removeItem(TEAM_EMAIL_KEY);
      } else {
        localStorage.setItem(SESSION_KEY, "team");
        localStorage.setItem(TEAM_TOKEN_KEY, payload.supabaseToken);
        localStorage.setItem(TEAM_EMAIL_KEY, payload.user?.email || "");
      }
    }
  };

  const handleLogout = () => {
    if (session?.mode === "google") signOut();
    localStorage.removeItem(SESSION_KEY);
    localStorage.removeItem(TEAM_TOKEN_KEY);
    localStorage.removeItem(TEAM_EMAIL_KEY);
    setSession(null);
    setLoggedIn(false);
    setProducts([]);
    setSyncStatus("idle");
  };

  // On mount: initialize Google Identity Services, then try to silently
  // restore whichever session type was last active — no re-login needed
  // unless the underlying token has actually expired or been revoked.
  //
  // Google's silent token request can, in some browser conditions, never
  // call back at all rather than erroring out — so we race it against a
  // timeout and treat "took too long" the same as "failed."
  const withTimeout = (promise, ms) =>
    Promise.race([promise, new Promise((resolve) => setTimeout(() => resolve(null), ms))]);

  useEffect(() => {
    let cancelled = false;

    // Absolute last-resort failsafe: no matter what hangs anywhere in the
    // chain above (script loading, silent auth, network), never leave the
    // person staring at a spinner forever. Worst case, they just see the
    // login screen a bit later than ideal instead of a stuck page.
    const hardCeiling = setTimeout(() => {
      if (!cancelled) setCheckingSession(false);
    }, 9000);

    initGoogleAuth(GOOGLE_CLIENT_ID).then(async () => {
      if (cancelled) return;
      setGoogleReady(true);

      const mode = localStorage.getItem(SESSION_KEY);

      if (mode === "google") {
        try {
          const restored = await withTimeout(trySilentSignIn(), 5000);
          if (!restored) throw new Error("Silent sign-in unavailable or timed out");
          const { sheetId, isNew } = await getOrCreateUserSheet(restored.token);
          const products = isNew ? [] : await readProducts(restored.token, sheetId);
          if (!cancelled) {
            handleSignedIn({ mode: "google", token: restored.token, user: restored.user, sheetId, products }, { persist: false });
          }
        } catch {
          localStorage.removeItem(SESSION_KEY);
        }
      } else if (mode === "team") {
        const token = localStorage.getItem(TEAM_TOKEN_KEY);
        const email = localStorage.getItem(TEAM_EMAIL_KEY);
        if (token) {
          try {
            const products = await readProductsViaBridge(token);
            if (!cancelled) {
              handleSignedIn({ mode: "team", supabaseToken: token, user: { email }, products }, { persist: false });
            }
          } catch {
            localStorage.removeItem(SESSION_KEY);
            localStorage.removeItem(TEAM_TOKEN_KEY);
            localStorage.removeItem(TEAM_EMAIL_KEY);
          }
        }
      }

      if (!cancelled) setCheckingSession(false);
    });

    return () => { cancelled = true; clearTimeout(hardCeiling); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Every mutation: update local state immediately (snappy UI), then push
  // the full list through to Google — directly if the owner, or via the
  // backend bridge if it's a team member — then fire a non-blocking backup.
  const syncToSheet = async (updatedProducts) => {
    if (!session) return;
    setSyncStatus("saving");
    try {
      if (session.mode === "google") {
        await writeAllProducts(session.token, session.sheetId, updatedProducts);
      } else {
        await writeProductsViaBridge(session.supabaseToken, updatedProducts);
      }
      setSyncStatus("saved");
      backupProducts(session.user?.email, updatedProducts); // fire-and-forget
      setTimeout(() => setSyncStatus((s) => (s === "saved" ? "idle" : s)), 2500);
    } catch (err) {
      console.error("Sheet sync failed:", err);
      setSyncStatus("error");
    }
  };

  // Human-readable timestamp for the sheet — this metadata is never shown
  // in the UI, it exists purely for reference when someone opens the sheet.
  const now = () => new Date().toLocaleString("en-IN", { dateStyle: "medium", timeStyle: "short" });

  const handleAdd = (product) => {
    const stamped = { ...product, createdAt: now(), updatedAt: now() };
    const updated = [stamped, ...products];
    setProducts(updated);
    syncToSheet(updated);
  };
  const handleEdit = (edited) => {
    const updated = products.map((p) =>
      p.id === edited.id ? { ...edited, createdAt: p.createdAt || now(), updatedAt: now() } : p
    );
    setProducts(updated);
    syncToSheet(updated);
  };
  const handleDelete = (id) => {
    const updated = products.filter((p) => p.id !== id);
    setProducts(updated);
    syncToSheet(updated);
  };
  const handleImport = (newProducts) => {
    const stamped = newProducts.map((p) => ({ ...p, createdAt: now(), updatedAt: now() }));
    const updated = [...stamped, ...products];
    setProducts(updated);
    syncToSheet(updated);
  };

  if (checkingSession) {
    return (
      <div style={{ fontFamily: "'Inter', sans-serif", background: COLORS.porcelain, height: "100vh" }} className="flex items-center justify-center">
        <style>{FONT_IMPORT}</style>
        <div className="flex items-center gap-2.5" style={{ color: COLORS.mist }}>
          <RefreshCw size={16} className="animate-spin" />
          <span className="text-sm">Restoring your session...</span>
        </div>
      </div>
    );
  }

  if (!loggedIn) return <LoginScreen onSignedIn={handleSignedIn} googleReady={googleReady} />;

  return (
    <div
      style={{ fontFamily: "'Inter', sans-serif", color: COLORS.ink, background: COLORS.porcelain, height: "100vh", overflowY: "auto" }}
    >
      <style>{FONT_IMPORT}</style>
      <TopNav
        activeTab={activeTab}
        setActiveTab={(t) => setActiveTab(t)}
        onLogout={handleLogout}
        user={session?.user}
        sheetId={session?.sheetId}
        syncStatus={syncStatus}
        sessionMode={session?.mode}
      />
      {activeTab === "dashboard" ? (
        <DashboardTab products={products} onAdd={handleAdd} onEdit={handleEdit} onDelete={handleDelete} onImport={handleImport} goToProducts={() => setActiveTab("products")} />
      ) : (
        <ProductsTab products={products} onAdd={handleAdd} onEdit={handleEdit} onDelete={handleDelete} onImport={handleImport} />
      )}
    </div>
  );
}
