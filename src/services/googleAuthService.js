// services/googleAuthService.js
//
// Wraps Google Identity Services (GIS) for browser-based OAuth.
// No backend required — token lives in memory for the session.
// Scopes are intentionally narrow: drive.file only sees files THIS app
// creates, not the user's whole Drive.

const SCOPES = "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/userinfo.email https://www.googleapis.com/auth/userinfo.profile";

let tokenClient = null;
let accessToken = null;
let currentUser = null; // { email, name, picture }

// Call once, after the Google script has loaded (see index.html changes below).
// This always resolves within 6 seconds, even if the script fails to load —
// callers must handle tokenClient still being unset in that case (signIn/
// trySilentSignIn already do, by failing gracefully rather than hanging).
export function initGoogleAuth(clientId, onTokenReady) {
  return new Promise((resolve) => {
    let settled = false;
    const finish = () => {
      if (settled) return;
      settled = true;
      clearInterval(checkLoaded);
      clearTimeout(giveUp);
      resolve();
    };

    const checkLoaded = setInterval(() => {
      if (window.google?.accounts?.oauth2) {
        tokenClient = window.google.accounts.oauth2.initTokenClient({
          client_id: clientId,
          scope: SCOPES,
          callback: (response) => {
            if (response.error) {
              console.error("Google auth error:", response);
              return;
            }
            accessToken = response.access_token;
            onTokenReady?.(accessToken);
          },
        });
        finish();
      }
    }, 100);

    const giveUp = setTimeout(finish, 6000);
  });
}

// Triggers the Google sign-in popup. Returns a promise that resolves once
// we have both a token and basic profile info.
export function signIn() {
  return new Promise((resolve, reject) => {
    if (!tokenClient) return reject(new Error("Google auth not initialized"));
    tokenClient.callback = async (response) => {
      if (response.error) return reject(response);
      accessToken = response.access_token;
      try {
        const profile = await fetchProfile(accessToken);
        currentUser = profile;
        resolve({ token: accessToken, user: profile });
      } catch (err) {
        reject(err);
      }
    };
    tokenClient.requestAccessToken({ prompt: "consent" });
  });
}

// Silent re-auth — call this on app load to restore a session without
// showing the popup again, if the browser still trusts this app.
export function trySilentSignIn() {
  return new Promise((resolve) => {
    if (!tokenClient) return resolve(null);
    tokenClient.callback = async (response) => {
      if (response.error) return resolve(null);
      accessToken = response.access_token;
      try {
        const profile = await fetchProfile(accessToken);
        currentUser = profile;
        resolve({ token: accessToken, user: profile });
      } catch {
        resolve(null);
      }
    };
    tokenClient.requestAccessToken({ prompt: "" });
  });
}

async function fetchProfile(token) {
  const res = await fetch("https://www.googleapis.com/oauth2/v3/userinfo", {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) throw new Error("Failed to fetch Google profile");
  return res.json(); // { email, name, picture, ... }
}

export function signOut() {
  if (accessToken && window.google?.accounts?.oauth2) {
    window.google.accounts.oauth2.revoke(accessToken, () => {});
  }
  accessToken = null;
  currentUser = null;
}

export function getAccessToken() {
  return accessToken;
}

export function getCurrentUser() {
  return currentUser;
}
