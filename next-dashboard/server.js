/**
 * Standalone Node.js + Express.js Server for Sales Executive Dashboard
 * - 100% independent of Google Apps Script
 * - Built-in 1-click OAuth2 Google Login (/auth/google) using GOOGLE_CLIENT_ID & GOOGLE_CLIENT_SECRET
 * - Also supports Google Service Account credentials
 * - Serves pre-injected dashboard HTML from RAM in <15ms
 */
require('dotenv').config();
const fs = require('fs');
const path = require('path');
const express = require('express');
const compression = require('compression');
const {
  getCachedData,
  syncNow,
  getGoogleAuth,
  getOAuthConsentUrl,
  handleOAuthCallback,
  getHealthStatus
} = require('./lib/sheets-engine');

const PORT = Number(process.env.PORT) || 3000;
const app = express();

app.use(compression());
app.use(express.json());

// 1-Click OAuth2 Login Route (when using GOOGLE_CLIENT_ID + GOOGLE_CLIENT_SECRET)
app.get('/auth/google', (req, res) => {
  const url = getOAuthConsentUrl();
  if (!url) {
    return res.status(400).send('Please add GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET to next-dashboard/.env first, then restart the server.');
  }
  res.redirect(url);
});

// OAuth2 Callback Route - saves refresh token automatically to .tokens.json
app.get('/auth/google/callback', async (req, res) => {
  try {
    if (!req.query.code) throw new Error('No OAuth code received from Google');
    await handleOAuthCallback(String(req.query.code));
    res.redirect('/');
  } catch (err) {
    res.status(500).send('Google OAuth Error: ' + err.message);
  }
});

// Live Polling REST API (<5ms from RAM)
app.get('/api/leads', async (req, res) => {
  try {
    const data = await getCachedData(req.query.stamp || '');
    res.setHeader('Cache-Control', 'no-store');
    res.json(data);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Real-Time Health & Diagnostic Check
app.get('/api/health', (req, res) => {
  res.setHeader('Cache-Control', 'no-store');
  const health = getHealthStatus();
  res.status(health.status === 'healthy' ? 200 : 503).json(health);
});

// Instant Webhook Trigger (Allows Google Apps Script onEdit trigger to force sync in <1s)
app.all('/api/sync', async (req, res) => {
  try {
    await syncNow();
    res.setHeader('Cache-Control', 'no-store');
    res.json({ ok: true, health: getHealthStatus() });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

function resolveHtmlPath() {
  const candidates = [
    path.join(__dirname, '..', 'Index.html'),
    path.join(__dirname, 'Index.html'),
    path.resolve(process.cwd(), 'Index.html'),
    path.resolve(process.cwd(), '..', 'Index.html')
  ];
  for (const c of candidates) {
    if (fs.existsSync(c)) return c;
  }
  return path.join(__dirname, '..', 'Index.html');
}

// Main Executive Dashboard Route
app.get('/', async (req, res) => {
  try {
    const authState = getGoogleAuth();
    // If Client ID & Secret are set in .env, but the user hasn't clicked "Connect Google" yet:
    if (!authState.ready && authState.mode === 'oauth2_needs_consent') {
      return res.send(`<!DOCTYPE html><html><head><meta charset="utf-8"><title>Connect Google Sheet</title>
        <style>body{font:15px system-ui,sans-serif;background:#072a44;color:#fff;display:flex;align-items:center;justify-content:center;height:100vh;margin:0}
        .box{background:#fff;color:#0f172a;padding:32px;border-radius:20px;max-width:440px;text-align:center;box-shadow:0 20px 50px rgba(0,0,0,.35)}
        a.btn{display:inline-block;margin-top:16px;padding:12px 24px;background:#0284c7;color:#fff;text-decoration:none;font-weight:700;border-radius:12px}</style></head>
        <body><div class="box"><h2>Connect Google Sheets &amp; Drive</h2>
        <p>Click below once to authorize your Google account. The server will save your refresh token automatically so the dashboard runs 24/7 on its own.</p>
        <a class="btn" href="/auth/google">Sign in with Google &rarr;</a></div></body></html>`);
    }

    const htmlPath = resolveHtmlPath();
    let html = fs.readFileSync(htmlPath, 'utf8');
    const initData = await getCachedData('');
    const safeJson = JSON.stringify(initData).replace(/<\//g, '<\\/');
    html = html.replace('/*__INIT_DATA__*/', 'window.__INIT__ = ' + safeJson + ';');
    res.setHeader('Content-Type', 'text/html; charset=utf-8');
    res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
    res.setHeader('Pragma', 'no-cache');
    res.setHeader('Expires', '0');
    res.send(html);
  } catch (err) {
    res.status(500).send('Dashboard Error: ' + err.message);
  }
});

app.listen(PORT, () => {
  const authState = getGoogleAuth();
  console.log(`\n🚀 Sales Executive Dashboard running at: http://localhost:${PORT}`);
  if (authState.mode === 'oauth2_needs_consent') {
    console.log(`🔑 First-time OAuth setup: Open http://localhost:${PORT}/auth/google in your browser to connect your Google account.\n`);
  } else if (!authState.ready) {
    console.log(`⚠️  Add GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET to next-dashboard/.env and restart.\n`);
  } else {
    syncNow().catch(() => {});
  }
});
