# Standalone Node.js Sales Executive Dashboard

This app runs **100% independently on Node.js + Express.js** using the official Google Sheets API v4 & Google Drive API v3 (no Google Apps Script required).

---

## Step 1: What to Put in `next-dashboard/.env`

Open [next-dashboard/.env](./.env) and fill in your **Client ID** and **Client Secret**:

```env
# 1. Your Google Sheet ID
SPREADSHEET_ID=1CKMXX2oa7RNiN8c_1u3uGbXjHcxdba1SGu5Is_9iejw

# Optional: Only fetch specific tab(s), e.g. ONLY_TABS=Nov 2026,Oct 2026
# Leave empty to fetch all visible tabs
ONLY_TABS=

# 2. Your Google OAuth2 Client ID & Client Secret
GOOGLE_CLIENT_ID=your_client_id_here.apps.googleusercontent.com
GOOGLE_CLIENT_SECRET=your_client_secret_here
GOOGLE_REDIRECT_URI=http://localhost:3000/auth/google/callback

# Server Port & Background Refresh Interval (10s)
PORT=3000
SYNC_INTERVAL_MS=10000
```

### How to Create `GOOGLE_CLIENT_ID` & `GOOGLE_CLIENT_SECRET` (If you don't have them yet)
1. Go to [Google Cloud Console](https://console.cloud.google.com/).
2. Enable **Google Sheets API** and **Google Drive API** under *APIs & Services > Library*.
3. Go to *APIs & Services > Credentials > Create Credentials > OAuth client ID*.
4. Application type: **Web application**.
5. Under **Authorized redirect URIs**, add:
   ```
   http://localhost:3000/auth/google/callback
   ```
6. Copy the **Client ID** and **Client Secret** into [next-dashboard/.env](./.env).

---

## Step 2: How to Run the Node.js App

1. Install **Node.js (LTS)** from [https://nodejs.org](https://nodejs.org) if it isn't installed on your PC yet.
2. Open PowerShell or Command Prompt inside the `next-dashboard` folder:
   ```powershell
   cd "c:\Users\ADMIN\Documents\Ai base code manager\Parsi sheet script creator\next-dashboard"
   npm install
   npm start
   ```
3. Open **http://localhost:3000** in your browser.
4. On the very first launch, click **"Sign in with Google"** once.
   - The server automatically saves your OAuth refresh token to `.tokens.json` so it will run 24/7 on its own without ever asking you to sign in again!
