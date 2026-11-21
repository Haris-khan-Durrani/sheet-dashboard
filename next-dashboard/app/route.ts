import fs from 'fs';
import path from 'path';
import { NextResponse } from 'next/server';
import { getCachedData, getGoogleAuth } from '../lib/sheets-engine';

export const dynamic = 'force-dynamic';
export const revalidate = 0;

function resolveHtmlPath() {
  const candidates = [
    path.resolve(process.cwd(), '..', 'Index.html'),
    path.resolve(process.cwd(), 'Index.html'),
    path.join(__dirname, '..', '..', 'Index.html'),
    path.join(__dirname, '..', 'Index.html')
  ];
  for (const c of candidates) {
    if (fs.existsSync(c)) return c;
  }
  return path.resolve(process.cwd(), '..', 'Index.html');
}

export async function GET() {
  try {
    const authState = getGoogleAuth();
    if (!authState.ready && authState.mode === 'oauth2_needs_consent') {
      const loginHtml = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Connect Google Sheet</title>
        <style>body{font:15px system-ui,sans-serif;background:#072a44;color:#fff;display:flex;align-items:center;justify-content:center;height:100vh;margin:0}
        .box{background:#fff;color:#0f172a;padding:32px;border-radius:20px;max-width:440px;text-align:center;box-shadow:0 20px 50px rgba(0,0,0,.35)}
        a.btn{display:inline-block;margin-top:16px;padding:12px 24px;background:#0284c7;color:#fff;text-decoration:none;font-weight:700;border-radius:12px}</style></head>
        <body><div class="box"><h2>Connect Google Sheets &amp; Drive</h2>
        <p>Click below once to authorize your Google account. The server will save your refresh token automatically so the dashboard runs 24/7 on its own.</p>
        <a class="btn" href="/auth/google">Sign in with Google &rarr;</a></div></body></html>`;
      return new NextResponse(loginHtml, {
        headers: {
          'Content-Type': 'text/html; charset=utf-8',
          'Cache-Control': 'no-store, no-cache, must-revalidate, proxy-revalidate'
        }
      });
    }

    const htmlPath = resolveHtmlPath();
    let html = fs.readFileSync(htmlPath, 'utf8');
    const initData = await getCachedData('');
    const safeJson = JSON.stringify(initData).replace(/<\//g, '<\\/');
    html = html.replace('/*__INIT_DATA__*/', 'window.__INIT__ = ' + safeJson + ';');
    return new NextResponse(html, {
      headers: {
        'Content-Type': 'text/html; charset=utf-8',
        'Cache-Control': 'no-store, no-cache, must-revalidate, proxy-revalidate',
        'Pragma': 'no-cache',
        'Expires': '0'
      }
    });
  } catch (err: any) {
    return new NextResponse('Dashboard Error: ' + err.message, { status: 500 });
  }
}
