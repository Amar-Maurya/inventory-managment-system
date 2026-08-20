// api/_shared.js
//
// Shared helpers used by all three backend functions. Reads secrets from
// environment variables — set these in Vercel Project Settings → Environment
// Variables, never commit them:
//
//   SUPABASE_URL                 e.g. https://xxxxx.supabase.co
//   SUPABASE_SERVICE_ROLE_KEY    from Supabase Project Settings → API
//   GOOGLE_SERVICE_ACCOUNT_KEY   the full JSON key file contents, as one line

import { google } from "googleapis";

export function getSupabaseConfig() {
  const url = process.env.SUPABASE_URL;
  const serviceKey = process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (!url || !serviceKey) throw new Error("Supabase is not configured on the server yet.");
  return { url, serviceKey };
}

// Verifies the Supabase session token sent from the browser and returns
// the signed-in team member's user id + email.
export async function getTeamMemberFromToken(authHeader) {
  const { url, serviceKey } = getSupabaseConfig();
  const token = authHeader?.replace("Bearer ", "");
  if (!token) throw new Error("Missing session token.");

  const res = await fetch(`${url}/auth/v1/user`, {
    headers: { apikey: serviceKey, Authorization: `Bearer ${token}` },
  });
  if (!res.ok) throw new Error("Session expired or invalid — please sign in again.");
  const user = await res.json();
  return { userId: user.id, email: user.email };
}

// Looks up which sheet this team member is allowed to access.
export async function getSheetForUser(userId) {
  const { url, serviceKey } = getSupabaseConfig();
  const res = await fetch(`${url}/rest/v1/team_access?user_id=eq.${userId}&select=sheet_id,owner_email`, {
    headers: { apikey: serviceKey, Authorization: `Bearer ${serviceKey}` },
  });
  if (!res.ok) throw new Error("Couldn't look up team access.");
  const rows = await res.json();
  if (!rows.length) throw new Error("This account isn't linked to any inventory yet.");
  return rows[0]; // { sheet_id, owner_email }
}

// Google Sheets client authenticated as the service account (a robot
// identity — no human login involved).
export function getServiceAccountSheets() {
  const raw = process.env.GOOGLE_SERVICE_ACCOUNT_KEY;
  if (!raw) throw new Error("Google service account is not configured on the server yet.");
  const credentials = JSON.parse(raw);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive.file"],
  });
  return google.sheets({ version: "v4", auth });
}
