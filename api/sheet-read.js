// api/sheet-read.js
//
// Called by a signed-in team member to load inventory. Verifies their
// Supabase session, finds which sheet they're linked to, then reads it
// using the service account (they never touch Google directly).

import { getTeamMemberFromToken, getSheetForUser, getServiceAccountSheets } from "./_shared.js";

export default async function handler(req, res) {
  if (req.method !== "GET") return res.status(405).json({ error: "Method not allowed" });

  try {
    const { userId } = await getTeamMemberFromToken(req.headers.authorization);
    const { sheet_id } = await getSheetForUser(userId);

    const sheets = getServiceAccountSheets();
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: sheet_id,
      range: "Inventory!A2:J1000",
    });

    const rows = result.data.values || [];
    const products = rows.map((r) => ({
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

    return res.status(200).json(products);
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
