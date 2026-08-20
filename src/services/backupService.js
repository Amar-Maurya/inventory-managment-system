// services/backupService.js
//
// One-way backup snapshot, fired after every successful write to the
// Google Sheet. The Sheet stays the single source of truth — this is
// purely a safety net in case it's ever deleted or corrupted.
//
// This is a stub until Supabase credentials are added. It fails silently
// (logs a warning) so a missing backup never blocks the actual save.

const SUPABASE_URL = null; // e.g. "https://xxxxx.supabase.co"
const SUPABASE_ANON_KEY = null;

export async function backupProducts(userEmail, products) {
  if (!SUPABASE_URL || !SUPABASE_ANON_KEY) {
    console.warn("[backup] Supabase not configured yet — skipping backup snapshot.");
    return;
  }
  try {
    await fetch(`${SUPABASE_URL}/rest/v1/inventory_backups`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        apikey: SUPABASE_ANON_KEY,
        Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
        Prefer: "resolution=merge-duplicates",
      },
      body: JSON.stringify({
        user_email: userEmail,
        products_json: products,
        backed_up_at: new Date().toISOString(),
      }),
    });
  } catch (err) {
    console.warn("[backup] Snapshot failed, non-blocking:", err.message);
  }
}
