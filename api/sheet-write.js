// api/sheet-write.js
//
// Mirrors sheet-read.js but for saving changes — same auth + lookup,
// then overwrites the sheet with the submitted product list.

import { getTeamMemberFromToken, getSheetForUser, getServiceAccountSheets } from "./_shared.js";

export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  try {
    const { userId } = await getTeamMemberFromToken(req.headers.authorization);
    const { sheet_id } = await getSheetForUser(userId);
    const { products } = req.body;

    const rows = (products || []).map((p) => [
      String(p.id), p.name, p.category, p.qty, p.price, p.supplier, p.threshold, p.dateAdded,
      p.createdAt || "", p.updatedAt || "",
    ]);

    const sheets = getServiceAccountSheets();
    await sheets.spreadsheets.values.update({
      spreadsheetId: sheet_id,
      range: `Inventory!A2:J${rows.length + 1}`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: rows },
    });

    return res.status(200).json({ success: true });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
