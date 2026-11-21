import { NextResponse } from 'next/server';
import { syncNow, getHealthStatus } from '../../../lib/sheets-engine';

export const dynamic = 'force-dynamic';

export async function POST() {
  try {
    await syncNow();
    return NextResponse.json({ ok: true, health: getHealthStatus() }, { headers: { 'Cache-Control': 'no-store' } });
  } catch (err: any) {
    return NextResponse.json({ ok: false, error: err.message }, { status: 500 });
  }
}

export async function GET() {
  return POST();
}
