// Google OAuth Client ID — public identifier, safe to keep in source control.
export const GOOGLE_CLIENT_ID = "32551401982-bu109bncu6em21vsi7h5ttcdt2jggo8i.apps.googleusercontent.com";

// The service account email that reads/writes sheets on behalf of
// email+password team members (via the backend bridge). Every new sheet
// gets shared with this address automatically. Fill this in once you've
// created the service account in Google Cloud Console — it's just an
// email address, safe to commit, not a secret.
export const SERVICE_ACCOUNT_EMAIL = "groutline-bridge@groutline-inventory.iam.gserviceaccount.com";
