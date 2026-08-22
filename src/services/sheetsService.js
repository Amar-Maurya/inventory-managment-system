// services/sheetsService.js
//
// All product data lives in a Google Sheet named "GroutLine Inventory".
// This file: finds that sheet for the signed-in user (or creates it),
// and maps between sheet rows and the product objects the UI uses.

import { SERVICE_ACCOUNT_EMAIL } from "../config";

const SHEET_NAME = "GroutLine Inventory";
const HEADERS = ["ID", "Name", "Category", "Quantity", "Price", "Supplier", "Low Stock Threshold", "Date Added", "Created", "Updated"];

function authHeaders(token) {
  return { Authorization: `Bearer ${token}`, "Content-Type": "application/json" };
}

// Looks for an existing sheet this app created. drive.file scope means
// this search only sees files the app itself has touched — nothing else
// in the user's Drive.
export async function findExistingSheet(token) {
  const query = encodeURIComponent(`name='${SHEET_NAME}' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false`);
  const res = await fetch(`https://www.googleapis.com/drive/v3/files?q=${query}&fields=files(id,name)`, {
    headers: authHeaders(token),
  });
  if (!res.ok) throw new Error("Drive search failed");
  const data = await res.json();
  return data.files?.[0]?.id || null;
}

// Creates a new spreadsheet with the correct headers, returns its file ID.
export async function createSheet(token) {
  const res = await fetch("https://sheets.googleapis.com/v4/spreadsheets", {
    method: "POST",
    headers: authHeaders(token),
    body: JSON.stringify({
      properties: { title: SHEET_NAME },
      sheets: [{ properties: { title: "Inventory" }, data: [{ rowData: [{ values: HEADERS.map((h) => ({ userEnteredValue: { stringValue: h } })) }] }] }],
    }),
  });
  if (!res.ok) throw new Error("Sheet creation failed");
  const data = await res.json();
  return data.spreadsheetId;
}

// Reads all rows and maps them to product objects for the UI.
// createdAt/updatedAt are carried through as hidden metadata — the UI
// never displays them, they exist purely for reference inside the sheet.
export async function readProducts(token, sheetId) {
  const res = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/Inventory!A2:J1000`, {
    headers: authHeaders(token),
  });
  if (!res.ok) throw new Error("Sheet read failed");
  const data = await res.json();
  const rows = data.values || [];
  return rows.map((r) => ({
    id: r[0] || String(Date.now() + Math.random()),
    name: r[1] || "Untitled product",
    category: r[2] || "Tiles",
    qty: Number(r[3]) || 0,
    price: Number(r[4]) || 0,
    supplier: r[5] || "—",
    threshold: Number(r[6]) || 10,
    dateAdded: r[7] || new Date().toISOString().slice(0, 10),
    createdAt: r[8] || null,
    updatedAt: r[9] || null,
  }));
}

// Overwrites the entire sheet with the current product list. Simple and
// reliable for a single-admin tool — avoids row-diffing complexity.
// For larger datasets this would move to targeted row updates instead.
export async function writeAllProducts(token, sheetId, products) {
  // Clear the full range first — otherwise, if the list has gotten
  // shorter (a delete), old rows past the new end would be left behind
  // and reappear as ghost products the next time the sheet is read.
  await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/Inventory!A2:J1000:clear`, {
    method: "POST",
    headers: authHeaders(token),
    body: JSON.stringify({}),
  });

  const rows = products.map((p) => [
    String(p.id), p.name, p.category, p.qty, p.price, p.supplier, p.threshold, p.dateAdded,
    p.createdAt || "", p.updatedAt || "",
  ]);
  if (rows.length === 0) return { success: true }; // nothing left to write, clear was enough

  const res = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/Inventory!A2:J${rows.length + 1}?valueInputOption=USER_ENTERED`,
    {
      method: "PUT",
      headers: authHeaders(token),
      body: JSON.stringify({ values: rows }),
    }
  );
  if (!res.ok) throw new Error("Sheet write failed");
  return res.json();
}

// Grants the backend's service account write access to this specific
// file — needed so email+password team members (who have no Google
// OAuth token of their own) can reach the sheet through the bridge API.
// Safe no-op if the service account isn't configured yet.
export async function shareWithServiceAccount(token, sheetId) {
  if (!SERVICE_ACCOUNT_EMAIL) return;
  try {
    await fetch(`https://www.googleapis.com/drive/v3/files/${sheetId}/permissions`, {
      method: "POST",
      headers: authHeaders(token),
      body: JSON.stringify({ role: "writer", type: "user", emailAddress: SERVICE_ACCOUNT_EMAIL }),
    });
  } catch (err) {
    console.warn("Couldn't share sheet with service account (team access will be unavailable):", err.message);
  }
}

// Convenience: get-or-create in one call, used right after sign-in.
export async function getOrCreateUserSheet(token) {
  let sheetId = await findExistingSheet(token);
  let isNew = false;
  if (!sheetId) {
    sheetId = await createSheet(token);
    isNew = true;
    await shareWithServiceAccount(token, sheetId);
  }
  return { sheetId, isNew };
}

// Direct link so the user can open the sheet in Google Sheets itself.
export function sheetUrl(sheetId) {
  return `https://docs.google.com/spreadsheets/d/${sheetId}/edit`;
}
