/**
 * Standalone Node.js Google Sheets + Drive API v3 Cell Comments Engine
 * - Uses globalThis.__DOCMAN_STATE__ so Next.js App Router routes share a single RAM cache (<5ms)
 * - Uses per-tab &gid= .xlsx export instead of drive.files.export (fixes "This file is too large to be exported")
 * - Non-blocking background comment coordinate resolution so initial load is <2s and subsequent loads are <5ms
 */
const fs = require('fs');
const path = require('path');
const { google } = require('googleapis');
const AdmZip = require('adm-zip');

const TOKEN_PATH = path.join(process.cwd(), '.tokens.json');
const ENV_PATH = path.join(process.cwd(), '.env');
const MONTHS = { jan:0, feb:1, mar:2, apr:3, may:4, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11 };
const SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets.readonly',
  'https://www.googleapis.com/auth/drive.readonly'
];

// Shared singleton state across Next.js hot-reloads and separate route bundles
if (!globalThis.__DOCMAN_STATE__) {
  globalThis.__DOCMAN_STATE__ = {
    memoryCache: {
      book: process.env.SPREADSHEET_ID || '1CKMXX2oa7RNiN8c_1u3uGbXjHcxdba1SGu5Is_9iejw',
      stamp: '',
      sheets: [],
      cerr: '',
      updatedAt: 0
    },
    xlsxCommentCache: { key: '', cellMap: {} },
    syncing: false,
    xlsxSyncing: false,
    timerStarted: false
  };
}
const state = globalThis.__DOCMAN_STATE__;

function getConfig() {
  return {
    spreadsheetId: (process.env.SPREADSHEET_ID || '1CKMXX2oa7RNiN8c_1u3uGbXjHcxdba1SGu5Is_9iejw').trim(),
    onlyTabs: (process.env.ONLY_TABS || '').split(',').map(s => s.trim().toLowerCase()).filter(Boolean),
    ignoreTabs: (process.env.IGNORE_TABS || '').split(',').map(s => s.trim().toLowerCase()).filter(Boolean),
    syncIntervalMs: Number(process.env.SYNC_INTERVAL_MS) || 10000
  };
}

function getGoogleAuth() {
  const clientId = (process.env.GOOGLE_CLIENT_ID || '').trim();
  const clientSecret = (process.env.GOOGLE_CLIENT_SECRET || '').trim();
  const redirectUri = (process.env.GOOGLE_REDIRECT_URI || `http://localhost:${process.env.PORT || 3000}/auth/google/callback`).trim();

  if (clientId && clientSecret) {
    const oauth2Client = new google.auth.OAuth2(clientId, clientSecret, redirectUri);
    let refreshToken = (process.env.GOOGLE_REFRESH_TOKEN || '').trim();
    let savedTokens = null;
    if (fs.existsSync(TOKEN_PATH)) {
      try { savedTokens = JSON.parse(fs.readFileSync(TOKEN_PATH, 'utf8')); } catch (e) {}
    }
    if (savedTokens && (savedTokens.refresh_token || savedTokens.access_token)) {
      oauth2Client.setCredentials(savedTokens);
      return { auth: oauth2Client, mode: 'oauth2', ready: true };
    }
    if (refreshToken) {
      oauth2Client.setCredentials({ refresh_token: refreshToken });
      return { auth: oauth2Client, mode: 'oauth2', ready: true };
    }
    return { auth: oauth2Client, mode: 'oauth2_needs_consent', ready: false };
  }

  const clientEmail = (process.env.GOOGLE_CLIENT_EMAIL || '').trim();
  const privateKey = (process.env.GOOGLE_PRIVATE_KEY || '').replace(/\\n/g, '\n').trim();
  if (clientEmail && privateKey) {
    const jwt = new google.auth.JWT({ email: clientEmail, key: privateKey, scopes: SCOPES });
    return { auth: jwt, mode: 'service_account_env', ready: true };
  }

  const keyFile = path.resolve(process.cwd(), process.env.GOOGLE_APPLICATION_CREDENTIALS || './service-account.json');
  if (fs.existsSync(keyFile)) {
    const gAuth = new google.auth.GoogleAuth({ keyFile, scopes: SCOPES });
    return { auth: gAuth, mode: 'service_account_file', ready: true };
  }

  return { auth: null, mode: 'unconfigured', ready: false };
}

function getOAuthConsentUrl() {
  const { auth, mode } = getGoogleAuth();
  if (!auth || (mode !== 'oauth2' && mode !== 'oauth2_needs_consent')) return null;
  return auth.generateAuthUrl({
    access_type: 'offline',
    prompt: 'consent',
    scope: SCOPES
  });
}

async function handleOAuthCallback(code) {
  const { auth } = getGoogleAuth();
  if (!auth) throw new Error('GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET are not set in .env');
  const { tokens } = await auth.getToken(code);
  fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens, null, 2), 'utf8');
  if (tokens.refresh_token) {
    process.env.GOOGLE_REFRESH_TOKEN = tokens.refresh_token;
    try {
      if (fs.existsSync(ENV_PATH)) {
        let envTxt = fs.readFileSync(ENV_PATH, 'utf8');
        if (/^GOOGLE_REFRESH_TOKEN=.*/m.test(envTxt)) {
          envTxt = envTxt.replace(/^GOOGLE_REFRESH_TOKEN=.*/m, 'GOOGLE_REFRESH_TOKEN=' + tokens.refresh_token);
        } else {
          envTxt += '\nGOOGLE_REFRESH_TOKEN=' + tokens.refresh_token + '\n';
        }
        fs.writeFileSync(ENV_PATH, envTxt, 'utf8');
      }
    } catch (e) {}
  }
  await syncNow();
  return tokens;
}

/**
 * Serves dashboard data directly from shared RAM (<5ms).
 */
async function getCachedData(knownStamp = '') {
  const cfg = getConfig();
  const authState = getGoogleAuth();
  if (!authState.ready) {
    return {
      sheets: [],
      cerr: authState.mode === 'oauth2_needs_consent' ? 'NEEDS_OAUTH_LOGIN' : 'NEEDS_ENV_CREDENTIALS',
      stamp: '',
      book: cfg.spreadsheetId
    };
  }
  if (!state.memoryCache.stamp) {
    await syncNow();
  } else if (Date.now() - state.memoryCache.updatedAt > cfg.syncIntervalMs && !state.syncing) {
    syncNow().catch(() => {});
  }
  if (knownStamp && knownStamp === state.memoryCache.stamp) {
    return { same: true, stamp: state.memoryCache.stamp, book: state.memoryCache.book };
  }
  return {
    sheets: state.memoryCache.sheets,
    cerr: state.memoryCache.cerr,
    stamp: state.memoryCache.stamp,
    book: state.memoryCache.book
  };
}

async function syncNow() {
  if (state.syncing) {
    // If another route already started syncing, wait briefly for it to populate RAM
    while (state.syncing && !state.memoryCache.stamp) {
      await new Promise(r => setTimeout(r, 150));
    }
    return;
  }
  const authState = getGoogleAuth();
  if (!authState.ready) return;
  state.syncing = true;
  try {
    const cfg = getConfig();
    const drive = google.drive({ version: 'v3', auth: authState.auth });
    const sheetsApi = google.sheets({ version: 'v4', auth: authState.auth });

    // 1. Fast change-stamp check (~150ms)
    const [fileMeta, latestCm] = await Promise.all([
      drive.files.get({ fileId: cfg.spreadsheetId, fields: 'modifiedTime' }),
      drive.comments.list({
        fileId: cfg.spreadsheetId,
        pageSize: 5,
        includeDeleted: false,
        fields: 'comments(modifiedTime)'
      }).catch(() => ({ data: { comments: [] } }))
    ]);

    const latestCommentTime = ((latestCm.data && latestCm.data.comments) || [])
      .map(c => c.modifiedTime || '')
      .sort()
      .pop() || '';
    const baseStamp = `${fileMeta.data.modifiedTime}:${cfg.onlyTabs.join('|')}:${latestCommentTime}`;
    const hasCoords = Object.keys(state.xlsxCommentCache.cellMap || {}).length > 0;
    const newStamp = baseStamp + (hasCoords ? ':cm' : '');

    if (newStamp === state.memoryCache.stamp && state.memoryCache.sheets.length) {
      state.memoryCache.updatedAt = Date.now();
      return;
    }

    // 2. Read sheet tabs via Google Sheets API v4 + recent Drive API comments in parallel (~1.2s)
    const [rawSheets, commentsList] = await Promise.all([
      readSheetsViaApi(sheetsApi, cfg),
      fetchAllDriveComments(drive, cfg.spreadsheetId)
    ]);

    // 3. Attach any already-cached cell comment coordinates immediately
    const totalMsgs = commentsList.reduce((s, x) => s + 1 + (x.rp ? x.rp.length : 0), 0);
    const xmKey = `${cfg.spreadsheetId}:${commentsList.length}:${totalMsgs}`;
    if (state.xlsxCommentCache.cellMap && Object.keys(state.xlsxCommentCache.cellMap).length) {
      attachCellComments(rawSheets, state.xlsxCommentCache.cellMap, commentsList);
    }

    const tabMeta = rawSheets.map(sh => ({ name: sh.name, gid: sh.gid }));
    const cleanSheets = finalizeSheets(rawSheets);

    state.memoryCache = {
      book: cfg.spreadsheetId,
      stamp: newStamp,
      sheets: cleanSheets,
      cerr: '',
      updatedAt: Date.now()
    };

    // 4. If comment coordinates aren't cached for this comment count yet, fetch per-tab .xlsx in the background!
    if (commentsList.length > 0 && state.xlsxCommentCache.key !== xmKey && !state.xlsxSyncing) {
      refreshPerTabCommentCoordsInBackground(authState.auth, cfg, tabMeta, xmKey);
    }
  } catch (err) {
    console.error('[SheetsEngine] Sync error:', err.message);
    state.memoryCache.cerr = err.message;
  } finally {
    state.syncing = false;
  }
}

/**
 * Exports each tab individually via &gid=<sheetId> so we NEVER hit Google Drive's 10MB whole-file export limit!
 * Runs asynchronously in the background without blocking page loads.
 */
async function refreshPerTabCommentCoordsInBackground(authClient, cfg, tabMeta, xmKey) {
  if (state.xlsxSyncing) return;
  state.xlsxSyncing = true;
  try {
    const tokenRes = await authClient.getAccessToken();
    const token = typeof tokenRes === 'string' ? tokenRes : (tokenRes && tokenRes.token);
    if (!token) return;

    const mergedMap = Object.assign({}, state.xlsxCommentCache.cellMap || {});
    // Export up to the first 6 visible tabs one by one using ?format=xlsx&gid=<sheetId>
    for (const tab of tabMeta.slice(0, 6)) {
      if (tab.gid == null) continue;
      const url = `https://docs.google.com/spreadsheets/d/${cfg.spreadsheetId}/export?format=xlsx&gid=${tab.gid}`;
      const ctrl = new AbortController();
      const timer = setTimeout(() => ctrl.abort(), 12000);
      try {
        const res = await fetch(url, {
          headers: { Authorization: `Bearer ${token}` },
          signal: ctrl.signal
        });
        if (res.ok) {
          const buf = Buffer.from(await res.arrayBuffer());
          const parsed = parseXlsxCommentsFromBuffer(buf, tab.name);
          if (parsed[tab.name]) mergedMap[tab.name] = parsed[tab.name];
        }
      } catch (e) {
        // Skip tab if export timed out
      } finally {
        clearTimeout(timer);
      }
    }

    state.xlsxCommentCache = { key: xmKey, cellMap: mergedMap };
    // Force a fast re-attach on the next sync tick
    state.memoryCache.stamp = '';
    await syncNow();
  } catch (e) {
    console.warn('[SheetsEngine] Background comment coordinate sync warning:', e.message);
  } finally {
    state.xlsxSyncing = false;
  }
}

function finalizeSheets(sheets) {
  return sheets.map(sh => {
    const copyRows = sh.rows.map(r => {
      const item = Object.assign({}, r);
      delete item.rn;
      Object.keys(item).forEach(k => { if (item[k] === '' || item[k] === null) delete item[k]; });
      return item;
    });
    return { name: sh.name, fh: sh.fh, rows: copyRows };
  });
}

async function readSheetsViaApi(sheetsApi, cfg) {
  const meta = await sheetsApi.spreadsheets.get({
    spreadsheetId: cfg.spreadsheetId,
    fields: 'sheets(properties(sheetId,title,hidden))'
  });

  const targetTabs = [];
  ((meta.data && meta.data.sheets) || []).forEach(s => {
    const p = s.properties || {};
    const title = String(p.title || '');
    const norm = title.trim().toLowerCase();
    if (cfg.onlyTabs.length) {
      if (cfg.onlyTabs.includes(norm)) targetTabs.push({ title, gid: p.sheetId });
    } else {
      if (!p.hidden && !cfg.ignoreTabs.includes(norm)) targetTabs.push({ title, gid: p.sheetId });
    }
  });

  if (!targetTabs.length) return [];

  // Batch-read all target tabs in ONE single API call (cap at 2500 rows per tab for speed)
  const ranges = targetTabs.map(t => `'${t.title.replace(/'/g, "''")}'!A1:AF2500`);
  const batchRes = await sheetsApi.spreadsheets.values.batchGet({
    spreadsheetId: cfg.spreadsheetId,
    ranges,
    valueRenderOption: 'FORMATTED_VALUE'
  });

  const out = [];
  ((batchRes.data && batchRes.data.valueRanges) || []).forEach((vr, idx) => {
    const tab = targetTabs[idx];
    const grid = vr.values || [];
    if (grid.length < 2) return;
    const hdr = findHeader(grid);
    if (!hdr) return;

    const c = hdr.cols, rows = [], labels = hdr.labels;
    const fcols = [];
    if (c.remarks > -1) {
      for (let i = c.remarks + 1; i < labels.length; i++) if (String(labels[i]).trim()) fcols.push(i);
    } else {
      labels.forEach((x, i) => { if (/DAY\s*\d/i.test(x)) fcols.push(i); });
    }

    for (let r = hdr.row + 1; r < grid.length; r++) {
      const row = grid[r];
      if (!row || !row.length) continue;
      const iso = parseDate(row[c.date]);
      if (!iso) continue;
      const name_ = getCell(row, c.name);
      const status = getCell(row, c.status);
      if (!name_ && !status) continue;
      const qd = c.qdate > -1 ? parseDate(row[c.qdate]) : '';
      const f = []; let anyF = false;
      for (let k = 0; k < fcols.length; k++) {
        const v = getCell(row, fcols[k]).slice(0, 300);
        f.push(v);
        if (v) anyF = true;
      }
      const item = {
        rn: r + 1, d: iso, qd, t: getCell(row, c.time), src: getCell(row, c.source), c60: getCell(row, c.called),
        name: name_, nat: getCell(row, c.nation), mob: getCell(row, c.mobile), st: status,
        asg: getCell(row, c.assigned), fz: getCell(row, c.fz), act: getCell(row, c.activity),
        visa: getCell(row, c.visa), soon: getCell(row, c.soon), rem: getCell(row, c.remarks).slice(0, 300)
      };
      if (anyF) item.f = f;
      rows.push(item);
    }

    out.push({
      name: tab.title,
      gid: tab.gid,
      rows,
      fc: fcols,
      lb: labels.map(x => String(x).replace(/\s+/g, ' ').trim()),
      fh: fcols.map(i => String(labels[i]).replace(/\s+/g, ' ').trim())
    });
  });

  return out;
}

async function fetchAllDriveComments(drive, fileId) {
  const norm = x => String(x || '').replace(/\s+/g, ' ').trim();
  const out = [];
  let pageToken = undefined, guard = 0;
  try {
    do {
      const res = await drive.comments.list({
        fileId,
        pageSize: 100,
        pageToken,
        includeDeleted: false,
        fields: 'nextPageToken,comments(content,createdTime,resolved,deleted,author/displayName,quotedFileContent/value,replies(content,createdTime,deleted,author/displayName))'
      });
      ((res.data && res.data.comments) || []).forEach(c => {
        if (c.deleted || c.resolved) return;
        const cText = norm(c.content).slice(0, 500);
        const rps = (c.replies || []).filter(r => !r.deleted && norm(r.content)).map(r => ({
          a: r.author ? r.author.displayName : '',
          t: r.createdTime,
          c: norm(r.content).slice(0, 500)
        }));
        if (!cText && !rps.length) return;
        out.push({
          a: c.author ? c.author.displayName : '',
          t: c.createdTime,
          c: cText,
          q: norm(c.quotedFileContent && c.quotedFileContent.value).slice(0, 300),
          rp: rps
        });
      });
      pageToken = res.data.nextPageToken;
      guard++;
    } while (pageToken && guard < 5);
  } catch (e) {
    console.warn('[SheetsEngine] Drive comments warning:', e.message);
  }
  return out;
}

function parseXlsxCommentsFromBuffer(buffer, fallbackTabName) {
  const zip = new AdmZip(buffer);
  const files = {};
  zip.getEntries().forEach(entry => {
    const n = entry.entryName;
    if (/^xl\/(workbook\.xml|_rels\/workbook\.xml\.rels|worksheets\/_rels\/[^\/]+\.rels|comments[^\/]*\.xml)$/.test(n)) {
      files[n] = entry.getData().toString('utf8');
    }
  });
  const attr = (tag, name) => { const m = tag.match(new RegExp('[ :]' + name + '="([^"]*)"')); return m ? m[1] : ''; };
  const dec = x => x.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&').replace(/&#(\d+);/g, (_, n) => String.fromCharCode(+n));
  const rels = {}, out = {};
  (files['xl/_rels/workbook.xml.rels'] || '').replace(/<Relationship\b[^>]*>/g, t => { rels[attr(t, 'Id')] = attr(t, 'Target'); return t; });
  (files['xl/workbook.xml'] || '').replace(/<sheet\b[^>]*>/g, t => {
    const name = dec(attr(t, 'name')) || fallbackTabName, file = (rels[attr(t, 'id')] || '').split('/').pop();
    const relXml = files['xl/worksheets/_rels/' + file + '.rels'];
    if (!file || !relXml) return t;
    relXml.replace(/<Relationship\b[^>]*>/g, r => {
      if (!/\/comments$/.test(attr(r, 'Type'))) return r;
      const xml = files['xl/' + attr(r, 'Target').split('/').pop()];
      if (!xml) return r;
      const list = [];
      xml.replace(/<comment\b([^>]*)>([\s\S]*?)<\/comment>/g, (m, a, body) => {
        const parts = [];
        body.replace(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g, (mm, x) => { parts.push(dec(x)); return mm; });
        const raw = parts.join('');
        list.push({ ref: attr(' ' + a, 'ref'), text: raw, msgs: parseXlsxThread(raw) });
        return m;
      });
      out[name] = list;
      if (fallbackTabName && !out[fallbackTabName]) out[fallbackTabName] = list;
      return r;
    });
    return t;
  });
  return out;
}

function parseXlsxThread(rawText) {
  const blocks = String(rawText || '').replace(/\r\n/g, '\n').split(/\n{2,}/);
  const msgs = [];
  blocks.forEach(b => {
    const trimmed = b.trim();
    if (!trimmed) return;
    const m = trimmed.match(/^([\s\S]*?)\n\s*-([^\n]+)$/);
    if (m) msgs.push({ a: m[2].trim(), t: '', c: m[1].replace(/\s+/g, ' ').trim() });
    else msgs.push({ a: '', t: '', c: trimmed.replace(/\s+/g, ' ').trim() });
  });
  return msgs;
}

function attachCellComments(sheets, cellMap, list) {
  const norm = x => String(x || '').replace(/\s+/g, ' ').trim();
  const cells = [];
  sheets.forEach(sh => {
    const byRn = {}; sh.rows.forEach(r => { byRn[r.rn] = r; });
    (cellMap[sh.name] || []).forEach(c => {
      const m = String(c.ref).match(/^([A-Z]+)(\d+)$/); if (!m) return;
      let col = 0; m[1].split('').forEach(ch => { col = col * 26 + ch.charCodeAt(0) - 64; });
      const row = byRn[+m[2]]; if (!row) return;
      const msgs = c.msgs || [];
      const rowMs = row.d ? new Date(row.d + 'T00:00:00Z').getTime() : 0;
      cells.push({
        sh, row, col: col - 1, rowMs,
        text: norm(c.text),
        msgs,
        sig: msgs.map(x => x.c.toLowerCase()).join('\x1f'),
        items: [],
        matched: false
      });
    });
  });

  const threads = (list || []).map(th => {
    const parts = [th.c].concat((th.rp || []).map(x => x.c)).filter(Boolean);
    const thMs = th.t ? new Date(th.t).getTime() : 0;
    return { th, thMs, parts, sig: parts.map(p => p.toLowerCase()).join('\x1f'), used: false };
  }).filter(x => x.parts.length > 0);

  const pickBestThread = (c, pred) => {
    let best = null, bestScore = Infinity;
    for (let i = 0; i < threads.length; i++) {
      const t = threads[i];
      if (t.used || !pred(t)) continue;
      const diffDays = (c.rowMs && t.thMs) ? (t.thMs - c.rowMs) / 86400000 : 0;
      const score = diffDays < -1 ? Math.abs(diffDays) * 1000 : Math.abs(diffDays);
      if (score < bestScore) { bestScore = score; best = t; }
    }
    return best;
  };

  cells.slice().sort((a, b) => b.msgs.length - a.msgs.length).forEach(c => {
    if (!c.sig) return;
    const best = pickBestThread(c, t => t.sig === c.sig);
    if (best) {
      c.items.push({ a: best.th.a, t: best.th.t, c: best.th.c, rp: best.th.rp });
      c.matched = true; best.used = true;
    }
  });

  cells.forEach(c => {
    if (c.matched) return;
    const best = pickBestThread(c, t => t.parts.every(p => c.text.indexOf(p) > -1));
    if (best) {
      c.items.push({ a: best.th.a, t: best.th.t, c: best.th.c, rp: best.th.rp });
      c.matched = true; best.used = true;
    }
  });

  cells.forEach(c => {
    let items = c.items;
    if (!items.length && c.msgs && c.msgs.length) items = [{ a: c.msgs[0].a, t: '', c: c.msgs[0].c, rp: c.msgs.slice(1) }];
    else if (!items.length && c.text) items = [{ a: '', t: '', c: c.text, rp: [] }];
    if (!items.length) return;
    const pos = c.sh.fc.indexOf(c.col);
    if (pos > -1) {
      c.row.cm = c.row.cm || {};
      c.row.cm[pos] = (c.row.cm[pos] || []).concat(items);
    } else {
      c.row.co = c.row.co || [];
      c.row.co.push({ h: c.sh.lb[c.col] || '', i: items });
    }
  });
}

function findHeader(grid) {
  const limit = Math.min(grid.length, 40);
  for (let r = 0; r < limit; r++) {
    const h = grid[r].map(x => String(x).toUpperCase().replace(/\s+/g, ' ').trim());
    const find = fn => { for (let i = 0; i < h.length; i++) if (h[i] && fn(h[i])) return i; return -1; };
    const date = find(x => x === 'DATE');
    const status = find(x => x.includes('LEAD STATUS') || x === 'STATUS');
    const name = find(x => x.includes('CLIENT NAME') || x === 'NAME');
    if (date > -1 && status > -1 && name > -1) {
      return {
        row: r,
        labels: grid[r],
        cols: {
          date, status, name,
          qdate: find(x => x.includes('QUALIFIED DATE') || x.includes('CLOSED DATE')),
          time: find(x => x === 'TIME'),
          source: find(x => x.includes('LEAD SOURCE')),
          called: find(x => x.includes('CALLED WITHIN')),
          nation: find(x => x.includes('NATION')),
          mobile: find(x => x.includes('MOBILE')),
          assigned: find(x => x.includes('ASSIGNED TO')),
          fz: find(x => x.includes('FREEZONE') || x.includes('FREE ZONE')),
          activity: find(x => x.includes('BUSINESS ACTIVITY')),
          visa: find(x => x.includes('VISA')),
          soon: find(x => x.includes('HOW SOON')),
          remarks: find(x => x.includes('REMARK'))
        }
      };
    }
  }
  return null;
}

function getCell(row, i) { return i > -1 && row[i] != null ? String(row[i]).replace(/\s+/g, ' ').trim() : ''; }

function parseDate(v) {
  const s = String(v || '').trim();
  let m = s.match(/^(\d{1,2})[-\/ ]([A-Za-z]{3,})[-\/ ,]*(\d{2,4})/);
  if (m) {
    const mo = MONTHS[m[2].slice(0, 3).toLowerCase()];
    if (mo === undefined) return '';
    const yr = +m[3] < 100 ? 2000 + +m[3] : +m[3];
    return fmtDate(yr, mo + 1, +m[1]);
  }
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return fmtDate(+m[1], +m[2], +m[3]);
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (m) {
    const yr = +m[3] < 100 ? 2000 + +m[3] : +m[3];
    return fmtDate(yr, +m[2], +m[1]);
  }
  return '';
}

function fmtDate(y, m, d) {
  if (y < 2000 || m < 1 || m > 12 || d < 1 || d > 31) return '';
  return y + '-' + ('0' + m).slice(-2) + '-' + ('0' + d).slice(-2);
}

if (!state.timerStarted) {
  state.timerStarted = true;
  setInterval(() => { syncNow().catch(() => {}); }, Number(process.env.SYNC_INTERVAL_MS) || 10000).unref();
}

module.exports = { getCachedData, syncNow, getGoogleAuth, getOAuthConsentUrl, handleOAuthCallback };
