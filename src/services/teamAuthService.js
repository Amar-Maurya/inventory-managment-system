// services/teamAuthService.js
//
// Handles email+password auth for TEAM MEMBERS (not the owner — the owner
// always uses Google). Built on Supabase Auth.
//
// Fill in SUPABASE_URL / SUPABASE_ANON_KEY once the project exists.
// Both are safe to keep in source — the anon key is public by design,
// Supabase enforces real security with Row Level Security policies, not
// by hiding this key.

const SUPABASE_URL = "https://kzaiidnqwovawmgvwyok.supabase.co";
const SUPABASE_ANON_KEY = "sb_publishable_8KbGiZhqomXX05r5Q7IsqA_0Aqtb_LV";

function ready() {
  return Boolean(SUPABASE_URL && SUPABASE_ANON_KEY);
}

// Called by the owner (while signed in with Google) to create a teammate
// login. Actual account creation needs the Supabase service-role key,
// which must stay server-side — so this calls our own backend function,
// not Supabase directly.
export async function inviteTeamMember({ ownerEmail, sheetId, teammateEmail, password }) {
  if (!ready()) throw new Error("Team sharing isn't configured yet.");
  const res = await fetch("/api/invite-team-member", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ownerEmail, sheetId, teammateEmail, password }),
  });
  if (!res.ok) throw new Error((await res.json())?.error || "Invite failed");
  return res.json();
}

// Teammate login — this part talks to Supabase directly, it's just a
// standard auth call, no secret key involved.
export async function signInWithEmail(email, password) {
  if (!ready()) throw new Error("Team sharing isn't configured yet.");
  const res = await fetch(`${SUPABASE_URL}/auth/v1/token?grant_type=password`, {
    method: "POST",
    headers: { "Content-Type": "application/json", apikey: SUPABASE_ANON_KEY },
    body: JSON.stringify({ email, password }),
  });
  if (!res.ok) throw new Error("Incorrect email or password.");
  const data = await res.json();
  // data.access_token is the Supabase session token — sent with every
  // request to our backend bridge so it knows which sheet to read/write.
  return { supabaseToken: data.access_token, email };
}

export function signOutTeamMember() {
  // Supabase sessions are stateless JWTs here — signing out client-side
  // is enough for this simple flow (no refresh-token persistence yet).
}
