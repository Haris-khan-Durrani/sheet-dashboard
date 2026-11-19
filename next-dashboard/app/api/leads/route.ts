import { NextResponse } from 'next/server';
import { getCachedData } from '../../../lib/sheets-engine';

export const dynamic = 'force-dynamic';

export async function GET(request: Request) {
  const { searchParams } = new URL(request.url);
  const stamp = searchParams.get('stamp') || '';
  const data = await getCachedData(stamp);
  return NextResponse.json(data);
}
