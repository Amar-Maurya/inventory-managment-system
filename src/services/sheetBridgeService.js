// services/sheetBridgeService.js
//
// Used only by email+password team members. They have no Google OAuth
// token, so instead of calling Sheets API directly (like the owner does),
// they call our own backend, which does it on their behalf using the
// service account.

export async function readProductsViaBridge(supabaseToken) {
  const res = await fetch("/api/sheet-read", {
    headers: { Authorization: `Bearer ${supabaseToken}` },
  });
  if (!res.ok) throw new Error((await res.json())?.error || "Couldn't load inventory.");
  return res.json();
}

export async function writeProductsViaBridge(supabaseToken, products) {
  const res = await fetch("/api/sheet-write", {
    method: "POST",
    headers: { "Content-Type": "application/json", Authorization: `Bearer ${supabaseToken}` },
    body: JSON.stringify({ products }),
  });
  if (!res.ok) throw new Error((await res.json())?.error || "Couldn't save changes.");
  return res.json();
}
