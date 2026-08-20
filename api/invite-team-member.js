// api/invite-team-member.js
//
// Called by the owner (while signed in with Google) to create a teammate
// login. Uses the Supabase service-role key, which only ever runs here,
// server-side — it's never sent to the browser.

import { getSupabaseConfig } from "./_shared.js";

export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  try {
    const { ownerEmail, sheetId, teammateEmail, password } = req.body;
    if (!ownerEmail || !sheetId || !teammateEmail || !password) {
      return res.status(400).json({ error: "Missing required fields." });
    }
    if (password.length < 8) {
      return res.status(400).json({ error: "Password must be at least 8 characters." });
    }

    const { url, serviceKey } = getSupabaseConfig();

    // 1. Create the auth user in Supabase.
    const createRes = await fetch(`${url}/auth/v1/admin/users`, {
      method: "POST",
      headers: { apikey: serviceKey, Authorization: `Bearer ${serviceKey}`, "Content-Type": "application/json" },
      body: JSON.stringify({ email: teammateEmail, password, email_confirm: true }),
    });
    if (!createRes.ok) {
      const err = await createRes.json();
      return res.status(400).json({ error: err.msg || "Couldn't create that login — email may already be in use." });
    }
    const newUser = await createRes.json();

    // 2. Link that user to the owner's sheet in a lookup table.
    // Create this table once in Supabase (SQL editor):
    //
    //   create table team_access (
    //     user_id uuid primary key,
    //     owner_email text not null,
    //     sheet_id text not null,
    //     created_at timestamptz default now()
    //   );
    const linkRes = await fetch(`${url}/rest/v1/team_access`, {
      method: "POST",
      headers: {
        apikey: serviceKey,
        Authorization: `Bearer ${serviceKey}`,
        "Content-Type": "application/json",
        Prefer: "return=minimal",
      },
      body: JSON.stringify({ user_id: newUser.id, owner_email: ownerEmail, sheet_id: sheetId }),
    });
    if (!linkRes.ok) {
      return res.status(500).json({ error: "Account created but couldn't link it to your sheet. Check the team_access table exists." });
    }

    return res.status(200).json({ success: true, email: teammateEmail });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
