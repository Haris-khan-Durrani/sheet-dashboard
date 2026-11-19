import { NextResponse } from 'next/server';
import { getOAuthConsentUrl } from '../../../lib/sheets-engine';

export const dynamic = 'force-dynamic';

export async function GET() {
  const url = getOAuthConsentUrl();
  if (!url) {
    return new NextResponse(
      'Please add GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET to next-dashboard/.env first, then restart the server.',
      { status: 400 }
    );
  }
  return NextResponse.redirect(url);
}
