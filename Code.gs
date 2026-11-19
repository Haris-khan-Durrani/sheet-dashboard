/**
 * Sales Dashboard - backend (with ONLY_TABS filter + Closed status + Qualified Date + mobile HTTP-0 proof)
 * Paste over your existing Code.gs. Index.html stays the file name.
 */
const SPREADSHEET_ID = '1CKMXX2oa7RNiN8c_1u3uGbXjHcxdba1SGu5Is_9iejw';

// Put the exact tab name(s) you want to fetch here, e.g. ['Nov 2026'] or ['Sheet1'].
// Leave empty [] to fetch all visible tabs (except IGNORE_TABS).
const ONLY_TABS = [];
const IGNORE_TABS = [];

const LIST_ALL_SPREADSHEETS = false;
const MONTHS = { jan:0, feb:1, mar:2, apr:3, may:4, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11 };

function doGet(e) {
  // JSON API mode for Node.js / Express.js / Next.js backend bridge (?json=1)
  if (e && e.parameter && e.parameter.json === '1') {
    const payload = getData(e.parameter.book || '', e.parameter.stamp || '');
    return ContentService.createTextOutput(JSON.stringify(payload))
      .setMimeType(ContentService.MimeType.JSON);
  }
  let html = HtmlService.createHtmlOutputFromFile('Index').getContent();
  try {
    // Inject initial data directly into the HTML response so mobile browsers render in 0ms without an XHR!
    const initData = getData('', '');
    const safeJson = JSON.stringify(initData).replace(/<\//g, '<\\/');
    html = html.replace('/*__INIT_DATA__*/', 'window.__INIT__ = ' + safeJson + ';');
  } catch (err) {
    // If anything fails during pre-injection, client falls back to localStorage + getData()
  }
  return HtmlService.createHtmlOutput(html)
    .setTitle('Sales Executive Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function boundId_() {
  return SPREADSHEET_ID || SpreadsheetApp.getActiveSpreadsheet().getId();
}

/** Spreadsheets the dashboard may show. Cached for 10 minutes. */
function getSpreadsheets() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('book-list');
  if (hit) return JSON.parse(hit);

  const homeId = boundId_(), seen = {}, list = [];
  const add = function (id, name) { if (!seen[id] && list.length < 100) { seen[id] = 1; list.push({ id: id, name: name }); } };
  add(homeId, DriveApp.getFileById(homeId).getName());

  const eat = function (it) { while (it.hasNext()) { const f = it.next(); add(f.getId(), f.getName()); } };
  if (LIST_ALL_SPREADSHEETS) {
    eat(DriveApp.getFilesByType(MimeType.GOOGLE_SHEETS));
  } else {
    const parents = DriveApp.getFileById(homeId).getParents();
    while (parents.hasNext()) eat(parents.next().getFilesByType(MimeType.GOOGLE_SHEETS));
  }
  list.sort(function (a, b) { return a.name.localeCompare(b.name); });
  cache.put('book-list', JSON.stringify(list), 600);
  return list;
}

/**
 * Fast change-stamp check:
 * Combines file lastUpdated + latest comment modifiedTime + ONLY_TABS config
 * and caches the stamp for 15 seconds so visitors get instant responses.
 */
function getFastStamp_(cache, bookId) {
  const tabKey = ONLY_TABS.join('|').toLowerCase();
  const stKey = 'st:' + bookId + ':' + tabKey;
  let stamp = cache.get(stKey);
  if (stamp) return stamp;
  stamp = String(DriveApp.getFileById(bookId).getLastUpdated().getTime()) + ':' + tabKey;
  if (typeof Drive !== 'undefined') {
    try {
      const res = Drive.Comments.list(bookId, { pageSize: 5, includeDeleted: false, fields: 'comments(modifiedTime)' });
      const latest = (res.comments || []).map(function (c) { return c.modifiedTime || ''; }).sort().pop() || '';
      stamp += ':' + latest;
    } catch (e) {
      stamp += ':' + Math.floor(Date.now() / 300000);
    }
  }
  cache.put(stKey, stamp, 15);
  return stamp;
}

function getData(id, knownStamp) {
  let bookId = boundId_();
  if (id && id !== bookId) {
    const ok = getSpreadsheets().some(function (b) { return b.id === id; });
    if (ok) bookId = id;
  }
  const cache = CacheService.getScriptCache();
  const stamp = getFastStamp_(cache, bookId);
  if (knownStamp && knownStamp === stamp) return { same: true, stamp: stamp, book: bookId };

  const key = 'd10:' + bookId + ':' + stamp;
  let payload = cacheGetBig_(cache, key);
  if (!payload) {
    payload = buildPayload_(cache, bookId);
    cachePutBig_(cache, key, payload);
  }
  return { sheets: payload.sheets, cerr: payload.cerr, stamp: stamp, book: bookId };
}

function buildPayload_(cache, bookId) {
  const sheets = readSheets_(bookId), cm = readComments_(bookId);
  let cells = {}, xerr = '';
  try {
    // Cache the expensive .xlsx export by comment count + reply count so sheet cell edits don't re-download .xlsx
    const totalMsgs = cm.list.reduce(function (s, x) { return s + 1 + (x.rp ? x.rp.length : 0); }, 0);
    const xmKey = 'xm3:' + bookId + ':' + cm.list.length + ':' + totalMsgs;
    cells = (cache && cacheGetBig_(cache, xmKey)) || null;
    if (!cells) {
      cells = readCellComments_(bookId);
      if (cache) cachePutBig_(cache, xmKey, cells);
    }
  } catch (e) {
    xerr = 'Could not read cell comments: ' + e.message;
    cells = {};
  }
  attachCellComments_(sheets, cells, cm.list);
  // Strip internal properties and empty string keys so mobile payload is ultra-small
  sheets.forEach(function (sh) {
    delete sh.fc; delete sh.lb;
    sh.rows.forEach(function (r) {
      delete r.rn;
      Object.keys(r).forEach(function (k) { if (r[k] === '' || r[k] === null) delete r[k]; });
    });
  });
  return { sheets: sheets, cerr: [cm.err, xerr].filter(Boolean).join(' ') };
}

/** Comments (with author + date/time + replies) from the Drive API v3 advanced service. */
function readComments_(bookId) {
  if (typeof Drive === 'undefined') return { list: [], err: 'Turn on comments: Apps Script > Services (+) > Drive API > Add, then deploy a new version.' };
  const norm = function (x) { return String(x || '').replace(/\s+/g, ' ').trim(); };
  try {
    const out = []; let token = null, guard = 0;
    do {
      const args = {
        pageSize: 100, includeDeleted: false,
        fields: 'nextPageToken,comments(content,createdTime,resolved,deleted,author/displayName,quotedFileContent/value,replies(content,createdTime,deleted,author/displayName))'
      };
      if (token) args.pageToken = token;
      const res = Drive.Comments.list(bookId, args);
      (res.comments || []).forEach(function (c) {
        if (c.deleted || c.resolved) return;
        const cText = norm(c.content).slice(0, 500);
        const rps = (c.replies || []).filter(function (r) { return !r.deleted && norm(r.content); }).map(function (r) {
          return { a: r.author ? r.author.displayName : '', t: r.createdTime, c: norm(r.content).slice(0, 500) };
        });
        if (!cText && !rps.length) return;
        out.push({
          a: c.author ? c.author.displayName : '', t: c.createdTime, c: cText,
          q: norm(c.quotedFileContent && c.quotedFileContent.value).slice(0, 300),
          rp: rps
        });
      });
      token = res.nextPageToken; guard++;
    } while (token && guard < 12);
    return { list: out };
  } catch (e) {
    return { list: [], err: 'Could not read comments: ' + e.message };
  }
}

/**
 * Google's .xlsx export lists every comment with its exact cell reference (e.g. "O12").
 */
function readCellComments_(bookId) {
  const urls = [
    'https://docs.google.com/spreadsheets/d/' + bookId + '/export?format=xlsx',
    'https://www.googleapis.com/drive/v3/files/' + bookId + '/export?mimeType=' + encodeURIComponent('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
  ];
  let res = null, last = '';
  for (let i = 0; i < urls.length && !res; i++) {
    const r = UrlFetchApp.fetch(urls[i], { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (r.getResponseCode() === 200) res = r;
    else last += (last ? ' | ' : '') + r.getResponseCode() + ' ' + r.getContentText().replace(/\s+/g, ' ').slice(0, 160);
  }
  if (!res) throw new Error('export failed: ' + last);
  const files = {};
  Utilities.unzip(res.getBlob().setContentType('application/zip')).forEach(function (b) {
    const n = b.getName();
    if (/^xl\/(workbook\.xml|_rels\/workbook\.xml\.rels|worksheets\/_rels\/[^\/]+\.rels|comments[^\/]*\.xml)$/.test(n)) {
      files[n] = b.getDataAsString();
    }
  });
  return parseCommentFiles_(files);
}

/** Parse Google Sheets .xlsx comment text into structured messages: [{a, t:'', c}, ...] */
function parseXlsxThread_(rawText) {
  const blocks = String(rawText || '').replace(/\r\n/g, '\n').split(/\n{2,}/);
  const msgs = [];
  blocks.forEach(function (b) {
    const trimmed = b.trim();
    if (!trimmed) return;
    const m = trimmed.match(/^([\s\S]*?)\n\s*-([^\n]+)$/);
    if (m) {
      msgs.push({ a: m[2].trim(), t: '', c: m[1].replace(/\s+/g, ' ').trim() });
    } else {
      msgs.push({ a: '', t: '', c: trimmed.replace(/\s+/g, ' ').trim() });
    }
  });
  return msgs;
}

function parseCommentFiles_(files) {
  const attr = function (tag, name) { const m = tag.match(new RegExp('[ :]' + name + '="([^"]*)"')); return m ? m[1] : ''; };
  const dec = function (x) { return x.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&').replace(/&#(\d+);/g, function (_, n) { return String.fromCharCode(+n); }); };
  const rels = {}, out = {};
  (files['xl/_rels/workbook.xml.rels'] || '').replace(/<Relationship\b[^>]*>/g, function (t) { rels[attr(t, 'Id')] = attr(t, 'Target'); return t; });
  (files['xl/workbook.xml'] || '').replace(/<sheet\b[^>]*>/g, function (t) {
    const name = dec(attr(t, 'name')), file = (rels[attr(t, 'id')] || '').split('/').pop();
    const relXml = files['xl/worksheets/_rels/' + file + '.rels'];
    if (!file || !relXml) return t;
    relXml.replace(/<Relationship\b[^>]*>/g, function (r) {
      if (!/\/comments$/.test(attr(r, 'Type'))) return r;
      const xml = files['xl/' + attr(r, 'Target').split('/').pop()];
      if (!xml) return r;
      const list = [];
      xml.replace(/<comment\b([^>]*)>([\s\S]*?)<\/comment>/g, function (m, a, body) {
        const parts = [];
        body.replace(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g, function (mm, x) { parts.push(dec(x)); return mm; });
        const raw = parts.join('');
        list.push({ ref: attr(' ' + a, 'ref'), text: raw, msgs: parseXlsxThread_(raw) });
        return m;
      });
      out[name] = list;
      return r;
    });
    return t;
  });
  return out;
}

/**
 * Matches Drive API comment threads (with exact timestamps & replies) to their exact cells.
 * Uses the row's lead date (r.d) so duplicate comments (e.g. "Not interested" on Nov 5 vs Nov 12)
 * always match the comment created on/after that row's lead date!
 */
function attachCellComments_(sheets, cellMap, list) {
  const norm = function (x) { return String(x || '').replace(/\s+/g, ' ').trim(); };
  const cells = [];
  sheets.forEach(function (sh) {
    const byRn = {}; sh.rows.forEach(function (r) { byRn[r.rn] = r; });
    (cellMap[sh.name] || []).forEach(function (c) {
      const m = String(c.ref).match(/^([A-Z]+)(\d+)$/); if (!m) return;
      let col = 0; m[1].split('').forEach(function (ch) { col = col * 26 + ch.charCodeAt(0) - 64; });
      const row = byRn[+m[2]]; if (!row) return;
      const msgs = c.msgs || [];
      const rowMs = row.d ? new Date(row.d + 'T00:00:00Z').getTime() : 0;
      cells.push({
        sh: sh, row: row, col: col - 1, rowMs: rowMs,
        text: norm(c.text),
        msgs: msgs,
        sig: msgs.map(function (x) { return x.c.toLowerCase(); }).join('\x1f'),
        items: [],
        matched: false
      });
    });
  });

  const threads = (list || []).map(function (th) {
    const parts = [th.c].concat((th.rp || []).map(function (x) { return x.c; })).filter(Boolean);
    const thMs = th.t ? new Date(th.t).getTime() : 0;
    return { th: th, thMs: thMs, parts: parts, sig: parts.map(function (p) { return p.toLowerCase(); }).join('\x1f'), used: false };
  }).filter(function (x) { return x.parts.length > 0; });

  const pickBestThread = function (c, pred) {
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

  cells.slice().sort(function (a, b) { return b.msgs.length - a.msgs.length; }).forEach(function (c) {
    if (!c.sig) return;
    const best = pickBestThread(c, function (t) { return t.sig === c.sig; });
    if (best) {
      c.items.push({ a: best.th.a, t: best.th.t, c: best.th.c, rp: best.th.rp });
      c.matched = true;
      best.used = true;
    }
  });

  cells.forEach(function (c) {
    if (c.matched) return;
    const best = pickBestThread(c, function (t) {
      return t.parts.every(function (p) { return c.text.indexOf(p) > -1; });
    });
    if (best) {
      c.items.push({ a: best.th.a, t: best.th.t, c: best.th.c, rp: best.th.rp });
      c.matched = true;
      best.used = true;
    }
  });

  cells.forEach(function (c) {
    let items = c.items;
    if (!items.length && c.msgs && c.msgs.length) {
      items = [{ a: c.msgs[0].a, t: '', c: c.msgs[0].c, rp: c.msgs.slice(1) }];
    } else if (!items.length && c.text) {
      items = [{ a: '', t: '', c: c.text, rp: [] }];
    }
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

/** Run from the editor to see what comments the script can find. View > Logs. */
function debugComments() {
  const id = boundId_(), cm = readComments_(id), cells = readCellComments_(id);
  Logger.log('Drive API comments: ' + cm.list.length + (cm.err ? ' (' + cm.err + ')' : ''));
  Object.keys(cells).forEach(function (n) {
    Logger.log('Tab "' + n + '": ' + cells[n].length + ' commented cells. First: ' + JSON.stringify(cells[n].slice(0, 3)));
  });
  if (!Object.keys(cells).length) Logger.log('No comments found in the Excel export of this file.');
}

/** Run once from the editor (Run > installWarmTrigger). Keeps the cache fresh so visitors rarely wait for a sheet read. */
function installWarmTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'warm') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('warm').timeBased().everyMinutes(5).create();
  warm();
}
function warm() {
  const id = boundId_();
  const tabKey = ONLY_TABS.join('|').toLowerCase();
  CacheService.getScriptCache().remove('st:' + id + ':' + tabKey);
  getData(id, '');
}

function cachePutBig_(cache, key, obj) {
  try {
    const b64 = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(JSON.stringify(obj), 'application/json')).getBytes());
    const SIZE = 90000, parts = Math.ceil(b64.length / SIZE), all = {};
    all[key + ':n'] = String(parts);
    for (let i = 0; i < parts; i++) all[key + ':' + i] = b64.substr(i * SIZE, SIZE);
    cache.putAll(all, 3600);
  } catch (e) { /* too big to cache: still works, just not shared */ }
}
function cacheGetBig_(cache, key) {
  try {
    const n = +cache.get(key + ':n');
    if (!n) return null;
    const keys = []; for (let i = 0; i < n; i++) keys.push(key + ':' + i);
    const got = cache.getAll(keys); let b64 = '';
    for (let i = 0; i < n; i++) { if (!got[keys[i]]) return null; b64 += got[keys[i]]; }
    const blob = Utilities.newBlob(Utilities.base64Decode(b64), 'application/x-gzip');
    return JSON.parse(Utilities.ungzip(blob).getDataAsString());
  } catch (e) { return null; }
}

function readSheets_(bookId) {
  const ss = SpreadsheetApp.openById(bookId);
  const only = ONLY_TABS.map(function (n) { return String(n).trim().toLowerCase(); }).filter(Boolean);
  const ignore = IGNORE_TABS.map(function (n) { return String(n).trim().toLowerCase(); }).filter(Boolean);
  const out = [];

  ss.getSheets().forEach(function (sh) {
    const name = sh.getName();
    const normName = name.trim().toLowerCase();
    // If ONLY_TABS is specified, skip every tab except the listed one(s)
    if (only.length) {
      if (only.indexOf(normName) === -1) return;
    } else {
      if (ignore.indexOf(normName) > -1 || sh.isSheetHidden()) return;
    }

    const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 2) return;

    const maxCol = Math.min(lastCol, 32);
    const grid = sh.getRange(1, 1, lastRow, maxCol).getDisplayValues();
    const hdr = findHeader_(grid);
    if (!hdr) return;

    const c = hdr.cols, rows = [], labels = hdr.labels;
    const fcols = [];
    if (c.remarks > -1) {
      for (let i = c.remarks + 1; i < labels.length; i++) if (String(labels[i]).trim()) fcols.push(i);
    } else {
      labels.forEach(function (x, i) { if (/DAY\s*\d/i.test(x)) fcols.push(i); });
    }

    for (let r = hdr.row + 1; r < grid.length; r++) {
      const row = grid[r];
      const iso = parseDate_(row[c.date]);
      if (!iso) continue;
      const name_ = get_(row, c.name);
      const status = get_(row, c.status);
      if (!name_ && !status) continue;
      const qd = c.qdate > -1 ? parseDate_(row[c.qdate]) : '';
      const f = []; let anyF = false;
      for (let k = 0; k < fcols.length; k++) {
        const v = get_(row, fcols[k]).slice(0, 300);
        f.push(v);
        if (v) anyF = true;
      }
      const item = {
        rn: r + 1, d: iso, qd: qd, t: get_(row, c.time), src: get_(row, c.source), c60: get_(row, c.called),
        name: name_, nat: get_(row, c.nation), mob: get_(row, c.mobile), st: status,
        asg: get_(row, c.assigned), fz: get_(row, c.fz), act: get_(row, c.activity),
        visa: get_(row, c.visa), soon: get_(row, c.soon), rem: get_(row, c.remarks).slice(0, 300)
      };
      if (anyF) item.f = f;
      rows.push(item);
    }
    out.push({
      name: name,
      rows: rows,
      fc: fcols,
      lb: labels.map(function (x) { return String(x).replace(/\s+/g, ' ').trim(); }),
      fh: fcols.map(function (i) { return String(labels[i]).replace(/\s+/g, ' ').trim(); })
    });
  });
  return out;
}

function findHeader_(grid) {
  const limit = Math.min(grid.length, 40);
  for (let r = 0; r < limit; r++) {
    const h = grid[r].map(function (x) { return String(x).toUpperCase().replace(/\s+/g, ' ').trim(); });
    const find = function (fn) { for (let i = 0; i < h.length; i++) if (h[i] && fn(h[i])) return i; return -1; };
    const date = find(function (x) { return x === 'DATE'; });
    const status = find(function (x) { return x.indexOf('LEAD STATUS') > -1 || x === 'STATUS'; });
    const name = find(function (x) { return x.indexOf('CLIENT NAME') > -1 || x === 'NAME'; });
    if (date > -1 && status > -1 && name > -1) {
      return {
        row: r,
        labels: grid[r],
        cols: {
          date: date, status: status, name: name,
          qdate: find(function (x) { return x.indexOf('QUALIFIED DATE') > -1 || x.indexOf('CLOSED DATE') > -1; }),
          time: find(function (x) { return x === 'TIME'; }),
          source: find(function (x) { return x.indexOf('LEAD SOURCE') > -1; }),
          called: find(function (x) { return x.indexOf('CALLED WITHIN') > -1; }),
          nation: find(function (x) { return x.indexOf('NATION') > -1; }),
          mobile: find(function (x) { return x.indexOf('MOBILE') > -1; }),
          assigned: find(function (x) { return x.indexOf('ASSIGNED TO') > -1; }),
          fz: find(function (x) { return x.indexOf('FREEZONE') > -1 || x.indexOf('FREE ZONE') > -1; }),
          activity: find(function (x) { return x.indexOf('BUSINESS ACTIVITY') > -1; }),
          visa: find(function (x) { return x.indexOf('VISA') > -1; }),
          soon: find(function (x) { return x.indexOf('HOW SOON') > -1; }),
          remarks: find(function (x) { return x.indexOf('REMARK') > -1; })
        }
      };
    }
  }
  return null;
}

function get_(row, i) { return i > -1 && row[i] != null ? String(row[i]).replace(/\s+/g, ' ').trim() : ''; }

function parseDate_(v) {
  const s = String(v || '').trim();
  let m = s.match(/^(\d{1,2})[-\/ ]([A-Za-z]{3,})[-\/ ,]*(\d{4})/);
  if (m) {
    const mo = MONTHS[m[2].slice(0, 3).toLowerCase()];
    if (mo === undefined) return '';
    return fmt_(+m[3], mo + 1, +m[1]);
  }
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return fmt_(+m[1], +m[2], +m[3]);
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return fmt_(+m[3], +m[2], +m[1]);
  return '';
}

function fmt_(y, m, d) {
  if (y < 2000 || m < 1 || m > 12 || d < 1 || d > 31) return '';
  return y + '-' + ('0' + m).slice(-2) + '-' + ('0' + d).slice(-2);
}
